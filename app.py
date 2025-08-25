import io, os, tempfile, subprocess, shutil
import streamlit as st
from dotenv import load_dotenv
from docx import Document
import fitz  # PyMuPDF

# ========= PPTX =========
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx_utils import translate_pptx_preserve_styles

# =================== Config & état ===================
st.set_page_config(page_title="Doc Translator", page_icon="🌐", layout="centered")
load_dotenv()

# Dossier de sortie local
os.makedirs("outputs", exist_ok=True)

def save_output_file(file_bytes, file_name):
    """Enregistre le fichier dans outputs/ et retourne le chemin."""
    path = os.path.join("outputs", file_name)
    with open(path, "wb") as f:
        f.write(file_bytes)
    return path

# État persistant
for k, v in {
    "translated_bytes": None,
    "translated_name": None,
    "translated_mime": None,
    "last_filename": None,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# =================== Helpers environnement & secrets ===================
def _get_secret(name, default=None):
    try:
        return st.secrets[name]
    except Exception:
        return os.getenv(name, default)

def get_openai_key():
    return _get_secret("OPENAI_API_KEY")

def has_ocr_binary():
    return shutil.which("ocrmypdf") is not None

def is_cloud_environment():
    """Heuristique simple pour Cloud : pas de binaire OCR ou désactivé via secret/env."""
    disabled = str(_get_secret("DISABLE_OCR", "0")) == "1"
    return disabled or not has_ocr_binary()

OCR_AVAILABLE_LOCALLY = has_ocr_binary()
RUNNING_IN_CLOUD = is_cloud_environment()
SHOW_OCR_BUTTON = OCR_AVAILABLE_LOCALLY and not RUNNING_IN_CLOUD  # Local OK, Cloud NON

# =================== Traduction (OpenAI si clé) ===================
def translate_batch(texts, src="fr", tgt="en"):
    """
    Traduit une liste de textes.
    - Si pas de clé: renvoie les textes d'origine (mode test)
    - Prompt orienté "sens" (pas mot-à-mot), ton neutre/pro.
    """
    api_key = get_openai_key()
    if not api_key:
        return texts

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        out = []
        for t in texts:
            system = (
                "You are a senior professional translator. Translate for MEANING and natural fluency, "
                "not word-by-word. Keep the original intent, register, and domain terminology. "
                "Preserve numbers, units, placeholders (like {name}), and punctuation. "
                "Do not add explanations."
            )
            user = f"Source language: {src}\nTarget language: {tgt}\n\nText:\n{t}"
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                ],
                temperature=0,
            )
            # Remplacer l’espace insécable par espace normal
            out.append(resp.choices[0].message.content.replace("\u00A0", " "))
        return out
    except Exception as e:
        st.error(f"Erreur API de traduction : {e}")
        return texts



    # =================== DOCX : traduction avancée (fluide + styles globaux) ===================
import re
from io import BytesIO
from docx import Document

# ---- Helpers : parsing UI (facultatif, fonctionne même si rien n'est défini dans l'UI) ----
def _parse_glossary_csv(csv_text: str) -> dict:
    """
    CSV simple: chaque ligne 'source,target'
    Insensible à la casse pour le repérage côté source.
    """
    d = {}
    if not csv_text:
        return d
    for line in csv_text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = [p.strip() for p in line.split(",")]
        if len(parts) >= 2 and parts[0] and parts[1]:
            d[parts[0]] = parts[1]
    return d

def _parse_dnt_terms(text: str) -> list:
    """
    Liste de termes à ne pas traduire. Séparés par ligne ou virgule.
    """
    if not text:
        return []
    items = []
    for chunk in re.split(r"[\n,]", text):
        t = chunk.strip()
        if t:
            items.append(t)
    # supprimer doublons en gardant l'ordre
    seen = set(); out = []
    for t in items:
        k = t.lower()
        if k not in seen:
            seen.add(k); out.append(t)
    return out

# ---- Helpers : protection/normalisation ----------------------------------------------------
def _normalize_text_after(t: str) -> str:
    # Espace insécable -> espace normal ; condense espaces multiples (hors sauts de ligne)
    t = t.replace("\u00A0", " ")
    # Remplacer séquences >1 espaces par 1, mais laisser \n/\r intacts
    t = re.sub(r"[ \t]{2,}", " ", t)
    return t

def _make_token(prefix: str, idx: int) -> str:
    # tokens sûrs qui ont peu de chances d’être inventés par le modèle
    return f"[[{prefix}{idx}]]"

def _protect_patterns(text: str) -> tuple[str, dict]:
    """
    Protège URLs, emails, {placeholders}, %s/%d, ACRONYMES.
    Retourne (texte_remplacé, mapping_token->valeur_originale)
    """
    mapping = {}
    idx = 0

    # Patterns
    patterns = [
        r"https?://\S+",
        r"[\w\.-]+@[\w\.-]+\.\w+",
        r"\{[^{}]+\}",            # {placeholder}
        r"%[sdif]",               # printf-like
        r"\b[A-Z]{2,}\b",         # ACRONYMES (2+ majuscules)
    ]

    def repl(m):
        nonlocal idx
        val = m.group(0)
        key = _make_token("TOK", idx); idx += 1
        mapping[key] = val
        return key

    for pat in patterns:
        text = re.sub(pat, repl, text)
    return text, mapping

def _protect_terms_ci(text: str, terms: list, prefix: str) -> tuple[str, dict]:
    """
    Remplace chaque terme (insensible à la casse) par un token unique.
    Retourne (texte, mapping token->valeur_originale)
    """
    mapping = {}
    if not terms:
        return text, mapping

    # trier par longueur (long d’abord) pour éviter les chevauchements
    terms_sorted = sorted(terms, key=lambda s: len(s), reverse=True)
    idx = 0

    for term in terms_sorted:
        # \b aux bords si le terme est "simple", sinon remplacement direct insensible casse
        t = re.escape(term)
        pat = rf"(?i){t}"
        def repl(m):
            nonlocal idx
            key = _make_token(prefix, idx); idx += 1
            mapping[key] = m.group(0)  # conserver le casing d’origine trouvé
            return key
        text = re.sub(pat, repl, text)
    return text, mapping

def _protect_glossary_ci(text: str, glossary: dict) -> tuple[str, dict]:
    """
    Pour le glossaire source→cible : remplace le terme source par un token [[GLOS#]].
    Post-traduction, [[GLOS#]] sera remplacé par la cible imposée.
    """
    if not glossary:
        return text, {}

    # trier par longueur (sources longues d’abord)
    items = sorted(glossary.items(), key=lambda kv: len(kv[0]), reverse=True)
    mapping = {}
    idx = 0
    for src, tgt in items:
        pat = rf"(?i){re.escape(src)}"
        def repl(m):
            nonlocal idx
            key = _make_token("GLOS", idx); idx += 1
            mapping[key] = tgt  # la cible fixée
            return key
        text = re.sub(pat, repl, text)
    return text, mapping

def _preprocess_text_for_translation(t: str, glossary: dict, dnt_terms: list) -> tuple[str, dict]:
    """
    Applique protections (TOK), DNT (KEEP) et Glossaire (GLOS).
    mapping global contient la priorité de restauration: GLOS -> KEEP -> TOK.
    """
    t1, map_tok  = _protect_patterns(t)
    t2, map_keep = _protect_terms_ci(t1, dnt_terms, "KEEP")
    t3, map_glos = _protect_glossary_ci(t2, glossary)
    # ordre de restau : GLOS -> KEEP -> TOK
    return t3, {"GLOS": map_glos, "KEEP": map_keep, "TOK": map_tok}

def _postprocess_translation(t: str, m: dict) -> str:
    # Restaurer dans l'ordre inverse : GLOS -> KEEP -> TOK
    for prefix in ("GLOS", "KEEP", "TOK"):
        for token, val in m.get(prefix, {}).items():
            t = t.replace(token, val)
    return _normalize_text_after(t)

# ---- DOCX traversals ----------------------------------------------------------------------
def _set_paragraph_text_preserve_para_style(p, new_text: str):
    """
    Remplace le contenu du paragraphe :
    - écrit tout dans le 1er run (il garde police/taille),
    - vide les suivants sans toucher au XML.
    """
    if p.runs:
        p.runs[0].text = new_text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(new_text)

def _collect_paragraph_objects_from_doc(doc: Document):
    """
    Collecte tous les paragraphes 'éditables' : corps, tableaux, en-têtes/pieds.
    Retourne une liste de références de paragraphes.
    """
    paras = []

    # Corps
    paras.extend(doc.paragraphs)

    # Tables du corps
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paras.extend(cell.paragraphs)

    # En-têtes / pieds
    for section in doc.sections:
        for part in (section.header, section.footer):
            paras.extend(part.paragraphs)
            for table in part.tables:
                for row in table.rows:
                    for cell in row.cells:
                        paras.extend(cell.paragraphs)
    return paras

# ---- Traduction principale pour DOCX -------------------------------------------------------
def translate_docx_preserve_styles(src_bytes, src="fr", tgt="en"):
    """
    Traduction DOCX 'au sens' avec :
    - Paragraphe par paragraphe (espaces corrects, meilleure qualité),
    - Corps + tableaux + en-têtes + pieds,
    - Glossaire (source→cible), Do-Not-Translate, protection tokens/URL/acronymes,
    - Normalisation espaces.
    Remarque: le micro-formatage intra-phrase (gras/italique partiels, hyperliens) peut être aplati.
    """

    # Lire options UI si présentes (sinon vide)
    glossary_csv = st.session_state.get("glossary_csv", "")  # ex: "serveur,server\nclient,customer"
    dnt_text     = st.session_state.get("dnt_terms", "")     # ex: "OpenAI\nGPU\nGPT-4o"

    GLOSSARY = _parse_glossary_csv(glossary_csv)
    DNT      = _parse_dnt_terms(dnt_text)

    doc = Document(BytesIO(src_bytes))

    # 1) Collecte des paragraphes
    paragraphs = _collect_paragraph_objects_from_doc(doc)

    # 2) Préparation batch (pré-process + collecte pour appel unique)
    para_refs = []
    to_translate = []
    preproc_maps = []

    for p in paragraphs:
        original = p.text or ""
        if original.strip():
            pre, maps = _preprocess_text_for_translation(original, GLOSSARY, DNT)
            para_refs.append(p)
            to_translate.append(pre)
            preproc_maps.append(maps)

    # 3) Traduction en batch (par paquets pour limiter taille des requêtes)
    translated_all = []
    BATCH = 50  # ajuste si besoin
    for i in range(0, len(to_translate), BATCH):
        chunk = to_translate[i:i+BATCH]
        out = translate_batch(chunk, src, tgt)  # réutilise ta fonction existante
        translated_all.extend(out)

    # 4) Post-process + écriture dans le doc
    for p, tr, maps in zip(para_refs, translated_all, preproc_maps):
        final_text = _postprocess_translation(tr, maps)
        _set_paragraph_text_preserve_para_style(p, final_text)

    # 5) Sauvegarde
    bio = BytesIO()
    doc.save(bio); bio.seek(0)
    return bio.read()


# =================== PDF : utilitaires ===================
def pdf_has_text(src_bytes, min_chars=20):
    """True si le PDF contient une couche texte suffisante."""
    try:
        doc = fitz.open(stream=src_bytes, filetype="pdf")
        has = False
        for page in doc:
            blocks = page.get_text("blocks")
            for b in blocks:
                if len(b) >= 5 and isinstance(b[4], str) and len(b[4].strip()) >= min_chars:
                    has = True
                    break
            if has:
                break
        doc.close()
        return has
    except Exception:
        return False

def ocr_pdf_with_ocrmypdf(src_bytes, lang="fra"):
    """
    OCR uniquement en local (désactivé en cloud).
    Pas de PDF/A ni d'optimisations lourdes. Timeout pour éviter les kills.
    """
    if RUNNING_IN_CLOUD:
        # Sécurité supplémentaire : ne jamais tenter l'OCR en cloud
        return src_bytes

    try:
        if pdf_has_text(src_bytes):
            return src_bytes

        with tempfile.TemporaryDirectory() as td:
            inp = os.path.join(td, "in.pdf")
            outp = os.path.join(td, "out.pdf")
            with open(inp, "wb") as f:
                f.write(src_bytes)
            cmd = [
                "ocrmypdf",
                "--skip-text",
                f"--language={lang}",
                "--output-type", "pdf",
                "--optimize", "0",
                "--fast-web-view", "0",
                inp, outp
            ]
            subprocess.run(
                cmd, check=True,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                timeout=120
            )
            with open(outp, "rb") as f:
                return f.read()
    except subprocess.TimeoutExpired:
        st.warning("⏱️ OCR trop long → on continue sans OCR.")
    except Exception as e:
        st.warning(f"OCR ignoré ({e})")
    return src_bytes

def translate_pdf_overlay(src_bytes, src="fr", tgt="en"):
    """
    Réécrit la couche texte traduite en overlay sur les blocs,
    avec ajustement automatique de la taille si besoin.
    """
    doc = fitz.open(stream=src_bytes, filetype="pdf")
    for page in doc:
        blocks = page.get_text("blocks")
        texts = [b[4] for b in blocks if len(b) >= 5 and isinstance(b[4], str) and b[4].strip()]
        if not texts:
            continue

        translated = translate_batch(texts, src, tgt)

        # Blanchir
        for (x0, y0, x1, y1, _txt, *_) in blocks:
            rect = fitz.Rect(x0, y0, x1, y1)
            page.add_redact_annot(rect, fill=(1, 1, 1))
        page.apply_redactions()

        # Réécrire avec ajustement
        def insert_text_fit(page, rect, text, fontname="Helvetica", max_size=11, min_size=6, step=0.5, align=0):
            size = max_size
            while size >= min_size:
                used = page.insert_textbox(rect, text, fontname=fontname, fontsize=size, align=align)
                if used >= 0:
                    return True
                page.add_redact_annot(rect, fill=(1, 1, 1))
                page.apply_redactions()
                size -= step
            page.insert_textbox(rect, text, fontname=fontname, fontsize=min_size, align=align)
            return False

        for (x0, y0, x1, y1, _txt, *_), new_text in zip(blocks, translated):
            rect = fitz.Rect(x0, y0, x1, y1)
            insert_text_fit(page, rect, new_text, max_size=11, min_size=6, step=0.5, align=0)

    out = io.BytesIO()
    doc.save(out)
    doc.close()
    out.seek(0)
    return out.read()

# =================== UI ===================
st.title("🌐 Document Translator (FR → EN) – format conservé")

src_lang = st.selectbox("Langue source", ["fr", "en", "es", "de"], index=0)
tgt_lang = st.selectbox("Langue cible", ["en", "fr", "es", "de"], index=1)

uploaded = st.file_uploader("Dépose ton fichier .docx, .pptx ou .pdf", type=["docx", "pptx", "pdf"])

if uploaded:
    data = uploaded.getvalue()
    st.info(f"Fichier reçu : {uploaded.name} ({len(data)} octets)")

    # Reset du résultat si fichier change
    if st.session_state.get("last_filename") != uploaded.name:
        st.session_state.translated_bytes = None
        st.session_state.translated_name = None
        st.session_state.translated_mime = None
        st.session_state.last_filename = uploaded.name

    name_lower = uploaded.name.lower()

    # ======== PDF ========
    if name_lower.endswith(".pdf"):
        if SHOW_OCR_BUTTON:
            if st.button("1) OCR (si scanné) → 2) Traduire PDF", key="btn_translate_pdf_ocr"):
                with st.spinner("Traitement PDF (OCR si besoin + traduction)…"):
                    try:
                        ocred = ocr_pdf_with_ocrmypdf(data, lang="fra" if src_lang == "fr" else src_lang)
                        translated = translate_pdf_overlay(ocred, src=src_lang, tgt=tgt_lang)
                        output_name = uploaded.name.replace(".pdf", f"_{tgt_lang}.pdf")

                        st.session_state.translated_bytes = translated
                        st.session_state.translated_name = output_name
                        st.session_state.translated_mime = "application/pdf"

                        save_path = save_output_file(translated, output_name)
                        st.success("✅ PDF traduit. Le bouton de téléchargement est prêt ci-dessous 👇")
                        st.info(f"💾 Fichier enregistré : {save_path}")
                    except Exception as e:
                        st.error(f"Erreur PDF/OCR: {e}")

            # Option locale: traduire sans OCR (si PDF déjà textuel)
            if st.button("Traduire PDF (sans OCR)", key="btn_translate_pdf_plain"):
                with st.spinner("Traduction PDF (sans OCR)…"):
                    try:
                        translated = translate_pdf_overlay(data, src=src_lang, tgt=tgt_lang)
                        output_name = uploaded.name.replace(".pdf", f"_{tgt_lang}.pdf")

                        st.session_state.translated_bytes = translated
                        st.session_state.translated_name = output_name
                        st.session_state.translated_mime = "application/pdf"

                        save_path = save_output_file(translated, output_name)
                        st.success("✅ PDF traduit (sans OCR).")
                        st.info(f"💾 Fichier enregistré : {save_path}")
                    except Exception as e:
                        st.error(f"Erreur PDF: {e}")

        else:
            st.warning("☁️ OCR désactivé en mode cloud (ou non disponible). "
                       "Les PDF scannés ne peuvent pas être convertis ici. "
                       "Traduction possible uniquement si le PDF contient déjà une couche texte.")
            if st.button("Traduire PDF (sans OCR)", key="btn_translate_pdf_cloud"):
                with st.spinner("Traduction PDF (sans OCR)…"):
                    try:
                        translated = translate_pdf_overlay(data, src=src_lang, tgt=tgt_lang)
                        output_name = uploaded.name.replace(".pdf", f"_{tgt_lang}.pdf")

                        st.session_state.translated_bytes = translated
                        st.session_state.translated_name = output_name
                        st.session_state.translated_mime = "application/pdf"

                        save_path = save_output_file(translated, output_name)
                        st.success("✅ PDF traduit (si le PDF était textuel).")
                        st.info(f"💾 Fichier enregistré : {save_path}")
                    except Exception as e:
                        st.error(f"Erreur PDF: {e}")

    # ======== DOCX ========
    elif name_lower.endswith(".docx"):
        if st.button("Traduire DOCX", key="btn_translate_docx"):
            with st.spinner("Traduction du DOCX en cours…"):
                try:
                    translated = translate_docx_preserve_styles(data, src=src_lang, tgt=tgt_lang)
                    output_name = uploaded.name.replace(".docx", f"_{tgt_lang}.docx")

                    st.session_state.translated_bytes = translated
                    st.session_state.translated_name = output_name
                    st.session_state.translated_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

                    save_path = save_output_file(translated, output_name)
                    st.success("✅ DOCX traduit. Le bouton de téléchargement est prêt ci-dessous 👇")
                    st.info(f"💾 Fichier enregistré : {save_path}")
                except Exception as e:
                    st.error(f"Erreur DOCX: {e}")

    # ======== PPTX ========
    elif name_lower.endswith(".pptx"):
        if st.button("Traduire PPTX", key="btn_translate_pptx"):
            with st.spinner("Traduction du PPTX en cours…"):
                try:
                    translated = translate_pptx_preserve_styles(
                        data, src=src_lang, tgt=tgt_lang, translate_callable=translate_batch
                    )
                    output_name = uploaded.name.replace(".pptx", f"_{tgt_lang}.pptx")

                    st.session_state.translated_bytes = translated
                    st.session_state.translated_name = output_name
                    st.session_state.translated_mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

                    save_path = save_output_file(translated, output_name)
                    st.success("✅ PPTX traduit. Le bouton de téléchargement est prêt ci-dessous 👇")
                    st.info(f"💾 Fichier enregistré : {save_path}")
                except Exception as e:
                    st.error(f"Erreur PPTX: {e}")

# Bouton de téléchargement commun
if st.session_state.translated_bytes:
    st.download_button(
        "⬇️ Télécharger le fichier traduit",
        data=st.session_state.translated_bytes,
        file_name=st.session_state.translated_name or "translated_file",
        mime=st.session_state.translated_mime or "application/octet-stream",
        key="download_translated_v1"
    )

st.divider()
st.write("⚙️ Conseils :")
st.write("- Ajoute ta clé dans `.env` ou dans les *Secrets* Streamlit Cloud (`OPENAI_API_KEY`).")
st.write("- PPTX : zones de texte, objets groupés, tableaux, titres/axes de graphiques pris en charge ; SmartArt/texte dans images non modifiables.")
st.write("- PDF : en Cloud, OCR désactivé. Les PDF scannés doivent être traités en local.")
