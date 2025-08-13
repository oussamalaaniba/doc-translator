import io, os, tempfile, subprocess, shutil, socket
import streamlit as st
from dotenv import load_dotenv
from docx import Document
import fitz  # PyMuPDF

# ========= PPTX =========
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

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
        if name in st.secrets:
            return st.secrets.get(name, default)
    except Exception:
        pass
    return os.getenv(name, default)

def get_openai_key():
    return _get_secret("OPENAI_API_KEY")

def has_ocr_binary():
    return shutil.which("ocrmypdf") is not None

def is_cloud_environment():
    """
    Heuristique simple pour Cloud : pas de binaire OCR ou désactivé via secret/env.
    """
    disabled = str(_get_secret("DISABLE_OCR", "0")) == "1"
    return disabled or not has_ocr_binary()

OCR_AVAILABLE_LOCALLY = has_ocr_binary()
RUNNING_IN_CLOUD = is_cloud_environment()
SHOW_OCR_BUTTON = OCR_AVAILABLE_LOCALLY and not RUNNING_IN_CLOUD  # Local OK, Cloud NON

# =================== Traduction (OpenAI si clé) ===================
def translate_batch(texts, src="fr", tgt="en"):
    """
    Traduit une liste de textes. Si pas de clé, renvoie les textes d'origine (mode écho).
    """
    api_key = get_openai_key()
    if not api_key:
        return texts  # mode test (pas de vraie traduction)

    try:
        from openai import OpenAI
        client = OpenAI()  # clé lue depuis env/secrets

        out = []
        for t in texts:
            prompt = (
                f"Translate from {src} to {tgt}. Preserve numbers, punctuation, line breaks, "
                f"and formatting hints. Output only the translation.\n\nTEXT:\n{t}"
            )
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are a professional translator."},
                    {"role": "user", "content": prompt},
                ],
                temperature=0,
            )
            out.append(resp.choices[0].message.content.strip())
        return out
    except Exception as e:
        st.error(f"Erreur API de traduction : {e}")
        return texts  # fallback

# =================== DOCX : préserver les styles ===================
def translate_docx_preserve_styles(src_bytes, src="fr", tgt="en"):
    doc = Document(io.BytesIO(src_bytes))
    runs_to_translate = []

    # Paragraphes
    for p in doc.paragraphs:
        for r in p.runs:
            if r.text.strip():
                runs_to_translate.append(r)

    # Tableaux
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.text.strip():
                            runs_to_translate.append(r)

    batch = [r.text for r in runs_to_translate]
    if batch:
        translated = translate_batch(batch, src, tgt)
        for r, new in zip(runs_to_translate, translated):
            r.text = new  # styles conservés

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.read()

# =================== PPTX : préserver styles & template ===================
def _collect_text_runs_from_shape(shape):
    """Retourne les runs (pptx.text.run.TextRun) d'une shape texte ou table."""
    runs = []
    if hasattr(shape, "has_text_frame") and shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.text.strip():
                    runs.append(r)
    elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
        table = shape.table
        for row in table.rows:
            for cell in row.cells:
                # Certaines cellules peuvent ne pas avoir text_frame
                if hasattr(cell, "text_frame") and cell.text_frame:
                    for p in cell.text_frame.paragraphs:
                        for r in p.runs:
                            if r.text.strip():
                                runs.append(r)
    return runs

def translate_pptx_preserve_styles(src_bytes, src="fr", tgt="en"):
    """
    - Text boxes, placeholders, tables -> traduction run-par-run (styles conservés)
    - Charts -> traduit titres et titres d'axes si accessibles
    - SmartArt/Diagrammes -> non supportés via python-pptx (texte inchangé)
    """
    prs = Presentation(io.BytesIO(src_bytes))
    run_refs = []

    for slide in prs.slides:
        for shape in slide.shapes:
            # 1) Text frames & tables
            run_refs.extend(_collect_text_runs_from_shape(shape))

            # 2) Charts: titre & axes (si dispo)
            if hasattr(shape, "has_chart") and shape.has_chart:
                chart = shape.chart
                # Titre
                try:
                    if chart.has_title and chart.chart_title.has_text_frame:
                        for p in chart.chart_title.text_frame.paragraphs:
                            for r in p.runs:
                                if r.text.strip():
                                    run_refs.append(r)
                except Exception:
                    pass
                # Axe catégories
                try:
                    if hasattr(chart, "category_axis") and chart.category_axis.has_title:
                        tf = chart.category_axis.axis_title.text_frame
                        for p in tf.paragraphs:
                            for r in p.runs:
                                if r.text.strip():
                                    run_refs.append(r)
                except Exception:
                    pass
                # Axe valeurs
                try:
                    if hasattr(chart, "value_axis") and chart.value_axis.has_title:
                        tf = chart.value_axis.axis_title.text_frame
                        for p in tf.paragraphs:
                            for r in p.runs:
                                if r.text.strip():
                                    run_refs.append(r)
                except Exception:
                    pass
                # NOTE: légendes/catégories/séries non éditées ici

    # Traduction
    batch = [r.text for r in run_refs]
    if batch:
        translated = translate_batch(batch, src, tgt)
        for r, new in zip(run_refs, translated):
            r.text = new

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
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
                    translated = translate_pptx_preserve_styles(data, src=src_lang, tgt=tgt_lang)
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
st.write("- PPTX : titres/axes de graphiques traduits ; SmartArt/diagrammes non modifiables via `python-pptx` (texte inchangé).")
st.write("- PDF : en Cloud, OCR désactivé. Les PDF *scannés* doivent être traités en local.")
