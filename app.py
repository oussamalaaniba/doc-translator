# app.py — Doc Translator (updated: strict JSON bulk + first‑page titles)
# - DOCX: bulk translation with JSON schema (reduces API calls massively)
# - Progress bar during DOCX
# - Includes headers + first-page/even-page headers + footers
# - NEW: translates text in Word text boxes (w:txbxContent), e.g., big titles on the cover page
# - OCR only when available locally; disabled in cloud by heuristic

import os, io, re, json, time, tempfile, subprocess, shutil
from io import BytesIO
from typing import List, Dict, Tuple

import streamlit as st
from dotenv import load_dotenv

# ===== Optional dependencies =====
try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover
    fitz = None

try:
    from docx import Document
    from docx.oxml.ns import nsmap, qn
    from lxml import etree as ET
except Exception:
    Document = None
    nsmap = {}
    def qn(x):
        return x  # type: ignore
    ET = None

# PPTX support (optional)
TRANSLATE_PPTX_AVAILABLE = True
try:
    from pptx import Presentation  # noqa: F401
    from pptx.enum.shapes import MSO_SHAPE_TYPE  # noqa: F401
    from pptx_utils import translate_pptx_preserve_styles
except Exception:
    TRANSLATE_PPTX_AVAILABLE = False
    translate_pptx_preserve_styles = None  # type: ignore

# =================== Config & état ===================
st.set_page_config(page_title="Doc Translator", page_icon="🌐", layout="centered")
load_dotenv()

# Dossier de sortie local
os.makedirs("outputs", exist_ok=True)

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

def _get_secret(name: str, default=None):
    try:
        return st.secrets[name]
    except Exception:
        return os.getenv(name, default)


def get_openai_key():
    # essaie plusieurs noms de clés pour compatibilité
    return (
        _get_secret("OPENAI_API_KEY")
        or _get_secret("openai_api_key")
        or _get_secret("OPENAI_KEY")
    )


def has_ocr_binary() -> bool:
    return shutil.which("ocrmypdf") is not None


def is_cloud_environment() -> bool:
    """Heuristique pour Cloud : OCR explicitement désactivé ou binaire absent."""
    disabled = str(_get_secret("DISABLE_OCR", "0")) == "1"
    return disabled or not has_ocr_binary()

OCR_AVAILABLE_LOCALLY = has_ocr_binary()
RUNNING_IN_CLOUD = is_cloud_environment()
SHOW_OCR_BUTTON = OCR_AVAILABLE_LOCALLY and not RUNNING_IN_CLOUD

# =================== Utilitaires IO ===================

def save_output_file(file_bytes: bytes, file_name: str) -> str:
    """Enregistre le fichier dans outputs/ et retourne le chemin."""
    path = os.path.join("outputs", file_name)
    with open(path, "wb") as f:
        f.write(file_bytes)
    return path

# =================== Traduction (OpenAI) ===================

def translate_batch(texts: List[str], src: str = "fr", tgt: str = "en", *, timeout: int = 60, max_retries: int = 3) -> List[str]:
    """
    Traduit une liste de textes en **UNE SEULE** requête (retour JSON) pour chaque batch.
    JSON schema strict + garde-fous pour maintenir la longueur.
    """
    n = len(texts)
    if n == 0:
        return []

    api_key = get_openai_key()
    if not api_key:
        # Pas de clé → mode dégradé : renvoie tel quel
        return texts

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        system = (
            "You are a professional translator. Translate EACH string in the JSON you receive "
            "from 'src' to 'tgt'. Preserve meaning, tone and punctuation. "
            "DO NOT translate placeholders or tokens like [[TOK#]], [[GLOS#]], [[KEEP#]], URLs or emails. "
            "Return ONLY JSON that matches the schema. No extra text."
        )
        user_payload = {"src": src, "tgt": tgt, "items": texts}
        user = json.dumps(user_payload, ensure_ascii=False)

        last_err = None
        for attempt in range(max_retries):
            try:
                kwargs = {
                    "model": "gpt-4o-mini",
                    "messages": [
                        {"role": "system", "content": system},
                        {"role": "user", "content": user},
                    ],
                    "temperature": 0,
                }
                # JSON schema strict (si supporté par la version du SDK)
                kwargs["response_format"] = {
                    "type": "json_schema",
                    "json_schema": {
                        "name": "batch_translations",
                        "strict": True,
                        "schema": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "items": {
                                    "type": "array",
                                    "minItems": n,
                                    "maxItems": n,
                                    "items": {"type": "string"}
                                }
                            },
                            "required": ["items"]
                        }
                    }
                }

                resp = client.chat.completions.create(**kwargs)
                content = resp.choices[0].message.content.strip()

                # Gérer les ```json ... ``` éventuels
                if content.startswith("```"):
                    content = content.strip("`")
                    content = re.sub(r"^json\n", "", content, flags=re.IGNORECASE)

                obj = json.loads(content)
                arr = obj["items"] if isinstance(obj, dict) and "items" in obj else obj

                if isinstance(arr, list) and len(arr) == n:
                    return [str(s).replace("\u00A0", " ") for s in arr]

                # Longueur inattendue → on retente
                raise ValueError(f"Bad JSON length: got {len(arr)} expected {n}")

            except Exception as e:
                last_err = e
                time.sleep(1.5 * (attempt + 1))

        # Fallback : item par item (lent) mais évite l'échec complet
        st.warning(f"Bulk translation failed ({last_err}); falling back to one-by-one.")
        out: List[str] = []
        for t in texts:
            out.extend(translate_batch([t], src, tgt, timeout=timeout, max_retries=1))
            time.sleep(0.05)
        return out
    except Exception as e:
        st.error(f"Erreur API de traduction : {e}")
        return texts

# =================== DOCX : helpers & pipeline ===================

ACRONYM_REGEX = r"\b[A-Z]{3,}\b"  # 3+ lettres pour éviter les faux positifs massifs


def _normalize_text_after(t: str) -> str:
    # Espace insécable -> espace normal ; condense espaces multiples (hors sauts de ligne)
    t = t.replace("\u00A0", " ")
    t = re.sub(r"[ \t]{2,}", " ", t)
    return t


def _make_token(prefix: str, idx: int) -> str:
    # tokens sûrs qui ont peu de chances d’être inventés par le modèle
    return f"[[{prefix}{idx}]]"


def _protect_patterns(text: str) -> Tuple[str, Dict[str, str]]:
    """
    Protège URLs, emails, {placeholders}, %s/%d, ACRONYMES.
    Retourne (texte_remplacé, mapping_token->valeur_originale)
    """
    mapping: Dict[str, str] = {}
    idx = 0

    patterns = [
        r"https?://\S+",
        r"[\w\.-]+@[\w\.-]+\.\w+",
        r"\{[^{}]+\}",        # {placeholder}
        r"%[sdif]",             # printf-like
        ACRONYM_REGEX,           # ACRONYMES (3+ majuscules)
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


def _protect_terms_ci(text: str, terms: List[str], prefix: str) -> Tuple[str, Dict[str, str]]:
    mapping: Dict[str, str] = {}
    if not terms:
        return text, mapping

    terms_sorted = sorted(terms, key=lambda s: len(s), reverse=True)
    idx = 0

    for term in terms_sorted:
        t = re.escape(term)
        pat = rf"(?i){t}"
        def repl(m):
            nonlocal idx
            key = _make_token(prefix, idx); idx += 1
            mapping[key] = m.group(0)  # conserver le casing d’origine trouvé
            return key
        text = re.sub(pat, repl, text)
    return text, mapping


def _protect_glossary_ci(text: str, glossary: Dict[str, str]) -> Tuple[str, Dict[str, str]]:
    if not glossary:
        return text, {}

    items = sorted(glossary.items(), key=lambda kv: len(kv[0]), reverse=True)
    mapping: Dict[str, str] = {}
    idx = 0
    for src, tgt in items:
        pat = rf"(?i){re.escape(src)}"
        def repl(m):
            nonlocal idx
            key = _make_token("GLOS", idx); idx += 1
            mapping[key] = tgt
            return key
        text = re.sub(pat, repl, text)
    return text, mapping


def _parse_glossary_csv(csv_text: str) -> Dict[str, str]:
    d: Dict[str, str] = {}
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


def _parse_dnt_terms(text: str) -> List[str]:
    if not text:
        return []
    items: List[str] = []
    for chunk in re.split(r"[\n,]", text):
        t = chunk.strip()
        if t:
            items.append(t)
    # supprimer doublons en gardant l'ordre
    seen = set(); out: List[str] = []
    for t in items:
        k = t.lower()
        if k not in seen:
            seen.add(k); out.append(t)
    return out
def _preprocess_text_for_translation(t: str, glossary: dict, dnt_terms: list) -> tuple[str, dict]:
    """
    Applique protections (TOK), DNT (KEEP) et Glossaire (GLOS).
    mapping global contient la priorité de restauration: GLOS -> KEEP -> TOK.
    """
    t1, map_tok  = _protect_patterns(t)
    t2, map_keep = _protect_terms_ci(t1, dnt_terms, "KEEP")
    t3, map_glos = _protect_glossary_ci(t2, glossary)
    return t3, {"GLOS": map_glos, "KEEP": map_keep, "TOK": map_tok}

def _postprocess_translation(t: str, m: dict) -> str:
    # Restaurer dans l'ordre inverse : GLOS -> KEEP -> TOK
    for prefix in ("GLOS", "KEEP", "TOK"):
        for token, val in m.get(prefix, {}).items():
            t = t.replace(token, val)
    # Espace insécable -> espace normal ; condense espaces multiples (hors sauts de ligne)
    t = t.replace("\u00A0", " ")
    t = re.sub(r"[ \t]{2,}", " ", t)
    return t
# ---- Paragraph setters -------------------------------------------------

def _set_docx_paragraph_text(p, new_text: str) -> None:
    """Remplace tout le contenu du paragraphe (python-docx) en gardant le style du 1er run."""
    if getattr(p, "runs", None) and p.runs:
        p.runs[0].text = new_text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(new_text)


def _set_xml_paragraph_text(p_xml, new_text: str) -> None:
    """Met à jour un <w:p> (paragraphe dans un textbox w:txbxContent) via XML."""
    if ET is None:
        return
    ts = p_xml.xpath('.//w:t', namespaces=nsmap) if nsmap else []
    if ts:
        ts[0].text = new_text
        for t in ts[1:]:
            t.text = ''
    else:
        # créer un run minimal r/t
        r = ET.Element(qn('w:r'))
        t = ET.SubElement(r, qn('w:t'))
        t.text = new_text
        p_xml.append(r)

# ---- Collecte des paragraphes -----------------------------------------

def _collect_txbx_paragraphs_from_element(root) -> List:
    """Renvoie la liste des <w:p> à l'intérieur des text boxes (w:txbxContent)."""
    if root is None or ET is None or not nsmap:
        return []
    return list(root.xpath('.//w:txbxContent//w:p', namespaces=nsmap))


def _collect_paragraph_objects_from_doc(doc) -> List[Tuple[str, object]]:
    """Retourne une liste de tuples (kind, ref), kind in {"docx", "xml"}.
    - "docx": objet Paragraph python-docx
    - "xml" : élément lxml <w:p> (ex: text boxes)
    """
    out: List[Tuple[str, object]] = []

    # Corps (Paragraph python-docx)
    for p in doc.paragraphs:
        out.append(("docx", p))

    # Tables du corps
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    out.append(("docx", p))

    # Text boxes du corps
    try:
        body_xml = doc.element.body  # lxml element
        for p_xml in _collect_txbx_paragraphs_from_element(body_xml):
            out.append(("xml", p_xml))
    except Exception:
        pass

    # En-têtes / pieds (incl. première page / pages paires si dispo)
    for section in doc.sections:
        for hdr_name in ("header", "first_page_header", "even_page_header"):
            part = getattr(section, hdr_name, None)
            if part is None:
                continue
            # Paragraphs
            for p in part.paragraphs:
                out.append(("docx", p))
            # Tables
            for table in part.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            out.append(("docx", p))
            # Text boxes dans le header
            try:
                root = part._element  # lxml element
                for p_xml in _collect_txbx_paragraphs_from_element(root):
                    out.append(("xml", p_xml))
            except Exception:
                pass

        for ftr_name in ("footer", "first_page_footer", "even_page_footer"):
            part = getattr(section, ftr_name, None)
            if part is None:
                continue
            for p in part.paragraphs:
                out.append(("docx", p))
            for table in part.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            out.append(("docx", p))
            try:
                root = part._element
                for p_xml in _collect_txbx_paragraphs_from_element(root):
                    out.append(("xml", p_xml))
            except Exception:
                pass

    return out

# ---- Pipeline ----------------------------------------------------------

def translate_docx_preserve_styles(src_bytes: bytes, src: str = "fr", tgt: str = "en") -> bytes:
    if Document is None:
        raise RuntimeError("python-docx manquant : impossible de traiter un DOCX.")

    # Lire options UI si présentes (sinon vide)
    glossary_csv = st.session_state.get("glossary_csv", "")
    dnt_text     = st.session_state.get("dnt_terms", "")

    GLOSSARY = _parse_glossary_csv(glossary_csv)
    DNT      = _parse_dnt_terms(dnt_text)

    doc = Document(BytesIO(src_bytes))

    # 1) Collecte de TOUTES les zones éditables (incl. text boxes & headers 1re page)
    collected = _collect_paragraph_objects_from_doc(doc)

    # 2) Préparation des entrées pour la traduction
    refs: List[Tuple[str, object]] = []
    to_translate: List[str] = []
    maps_list: List[Dict[str, Dict[str, str]]] = []

    for kind, ref in collected:
        # Récupération du texte courant
        if kind == "docx":
            original = getattr(ref, "text", "") or ""
        else:  # xml
            ts = ref.xpath('.//w:t', namespaces=nsmap) if nsmap else []
            original = "".join([t.text or "" for t in ts]) if ts else ""

        if original.strip():
            pre, maps = _preprocess_text_for_translation(original, GLOSSARY, DNT)
            refs.append((kind, ref))
            to_translate.append(pre)
            maps_list.append(maps)

    if not to_translate:
        # Rien à traduire → renvoyer l'original
        return src_bytes

    # 3) Traduction en chunk (BULK par chunk)
    translated_all: List[str] = []
    CHUNK = 18  # chunks plus petits = moins d'erreurs de longueur
    total = len(to_translate)
    prog = st.progress(0.0)

    for i in range(0, total, CHUNK):
        chunk = to_translate[i:i + CHUNK]
        out = translate_batch(chunk, src, tgt, timeout=60)
        # Si le modèle renvoie moins d'items, compléter par les originaux
        if len(out) < len(chunk):
            out = out + chunk[len(out):]
        translated_all.extend(out)
        prog.progress(min(1.0, len(translated_all) / total))
    prog.empty()

    # 4) Post-process + écriture
    for (kind, ref), tr, maps in zip(refs, translated_all, maps_list):
        final_text = _postprocess_translation(tr, maps)
        if kind == "docx":
            _set_docx_paragraph_text(ref, final_text)
        else:
            _set_xml_paragraph_text(ref, final_text)

    # 5) Sauvegarde
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.read()

# =================== PDF : utilitaires ===================

def pdf_has_text(src_bytes: bytes, min_chars: int = 20) -> bool:
    if fitz is None:
        return False
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


def ocr_pdf_with_ocrmypdf(src_bytes: bytes, lang: str = "fra") -> bytes:
    if RUNNING_IN_CLOUD:
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
                inp, outp,
            ]
            subprocess.run(
                cmd, check=True,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                timeout=120,
            )
            with open(outp, "rb") as f:
                return f.read()
    except subprocess.TimeoutExpired:
        st.warning("⏱️ OCR trop long → on continue sans OCR.")
    except Exception as e:
        st.warning(f"OCR ignoré ({e})")
    return src_bytes


def translate_pdf_overlay(src_bytes: bytes, src: str = "fr", tgt: str = "en") -> bytes:
    if fitz is None:
        raise RuntimeError("PyMuPDF manquant : la traduction PDF n'est pas disponible.")

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

st.title("🌐 Document Translator ")

src_lang = st.selectbox("Langue source", ["fr", "en", "es", "de"], index=0)
tgt_lang = st.selectbox("Langue cible", ["en", "fr", "es", "de"], index=1)

with st.expander("⚙️ Options traduction (DOCX)"):
    st.session_state["glossary_csv"] = st.text_area(
        "Glossaire source,target (CSV, une paire par ligne)",
        value=st.session_state.get("glossary_csv", ""),
        placeholder="serveur,server\nclient,customer",
    )
    st.session_state["dnt_terms"] = st.text_area(
        "Termes à NE PAS traduire (un par ligne ou séparés par des virgules)",
        value=st.session_state.get("dnt_terms", ""),
        placeholder="OpenAI\nGPU\nGPT-4o",
    )
    st.caption(
        "Astuce : le glossaire force une traduction précise de certains termes. "
        "Les termes à ne pas traduire (DNT) seront laissés tels quels."
    )

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
                        lang_ocr = "fra" if src_lang == "fr" else src_lang
                        ocred = ocr_pdf_with_ocrmypdf(data, lang=lang_ocr)
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
            st.warning(
                "☁️ OCR désactivé en mode cloud (ou non disponible). "
                "Les PDF scannés ne peuvent pas être convertis ici. "
                "Traduction possible uniquement si le PDF contient déjà une couche texte."
            )
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
        if not TRANSLATE_PPTX_AVAILABLE:
            st.error("Le module PPTX n'est pas disponible (pptx_utils manquant).")
        else:
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
        key="download_translated_v1",
    )

st.divider()
st.write("⚙️ Conseils :")
st.write("- Ajoute ta clé dans `.env` ou dans les *Secrets* Streamlit Cloud (`OPENAI_API_KEY`).")
st.write("- DOCX : désormais, les titres en page de garde (text boxes / en-tête 1re page) sont traduits.")
st.write("- PPTX : zones de texte, objets groupés, tableaux, titres/axes de graphiques pris en charge ; SmartArt/texte dans images non modifiables.")
st.write("- PDF : en Cloud, OCR désactivé. Les PDF scannés doivent être traités en local.")
