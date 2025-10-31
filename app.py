 (cd "$(git rev-parse --show-toplevel)" && git apply --3way <<'EOF' 
diff --git a/app.py b/app.py
index d1425325d5634e6c2c4f7436c756df23a40e5987..ac4a47e1c9e66ae04c936f6a359c124b49950469 100644
--- a/app.py
+++ b/app.py
@@ -1,62 +1,93 @@
 # app.py — Doc Translator + Background image (local & cloud-safe)
 import os, io, re, json, time, tempfile, subprocess, shutil, base64, zipfile
 from io import BytesIO
-from typing import List, Dict, Tuple, Any, Optional
+from typing import List, Dict, Tuple, Any, Optional, Callable
 from pathlib import Path
 
 import streamlit as st
 from dotenv import load_dotenv
 
 # ===== Optional deps =====
 try:
     import fitz  # PyMuPDF
 except Exception:
     fitz = None
 
 try:
     from docx import Document
     from docx.oxml.ns import qn
     from lxml import etree as ET
 except Exception:
     Document = None
     def qn(x): return x  # type: ignore
     ET = None
 
 # PPTX (optional)
 TRANSLATE_PPTX_AVAILABLE = True
 try:
     from pptx_utils import translate_pptx_preserve_styles  # noqa
 except Exception:
     TRANSLATE_PPTX_AVAILABLE = False
     translate_pptx_preserve_styles = None  # type: ignore
 
 # =================== Config ===================
 st.set_page_config(page_title="Doc Translator", page_icon="🌐", layout="centered")
 load_dotenv()
 os.makedirs("outputs", exist_ok=True)
 
+LANGUAGE_OPTIONS = [
+    ("Français", "fr"),
+    ("English", "en"),
+    ("Español", "es"),
+    ("Deutsch", "de"),
+    ("Italiano", "it"),
+    ("Português", "pt"),
+    ("Nederlands", "nl"),
+    ("Русский", "ru"),
+    ("العربية", "ar"),
+    ("中文 (简体)", "zh"),
+    ("日本語", "ja"),
+]
+
+LANGUAGE_LABELS = {code: label for label, code in LANGUAGE_OPTIONS}
+LANGUAGE_CODES = [code for _, code in LANGUAGE_OPTIONS]
+
+TESSERACT_LANG_MAP = {
+    "fr": "fra",
+    "en": "eng",
+    "es": "spa",
+    "de": "deu",
+    "it": "ita",
+    "pt": "por",
+    "nl": "nld",
+    "ru": "rus",
+    "ar": "ara",
+    "zh": "chi_sim",
+    "ja": "jpn",
+}
+
 for k, v in {"translated_bytes": None, "translated_name": None, "translated_mime": None, "last_filename": None}.items():
     if k not in st.session_state:
         st.session_state[k] = v
 
 # =================== Background (local or cloud) ===================
 def set_background_auto(default_filename: str = "bg.png", darken: float = 0.35):
     """
     Utilise bg.png à côté de app.py si présent. Sinon:
     - st.secrets['BACKGROUND_URL']  : URL https
     - st.secrets['BACKGROUND_BASE64']: base64 d'une image (sans 'data:')
 
     'darken' applique un voile sombre pour la lisibilité (0..0.6).
     """
     # 1) image locale (même dossier que app.py)
     here = Path(__file__).parent
     local = here / default_filename
     img_b64 = None
     url = None
 
     if local.exists():
         try:
             img_b64 = base64.b64encode(local.read_bytes()).decode()
         except Exception:
             img_b64 = None
 
@@ -688,183 +719,287 @@ def pdf_has_text(src_bytes: bytes, min_chars: int = 20) -> bool:
             blocks = page.get_text("blocks")
             for b in blocks:
                 if len(b) >= 5 and isinstance(b[4], str) and len(b[4].strip()) >= min_chars:
                     has = True; break
             if has: break
         doc.close(); return has
     except Exception:
         return False
 
 def ocr_pdf_with_ocrmypdf(src_bytes: bytes, lang: str = "fra") -> bytes:
     if RUNNING_IN_CLOUD: return src_bytes
     try:
         if pdf_has_text(src_bytes): return src_bytes
         with tempfile.TemporaryDirectory() as td:
             inp = os.path.join(td, "in.pdf"); outp = os.path.join(td, "out.pdf")
             with open(inp, "wb") as f: f.write(src_bytes)
             cmd = ["ocrmypdf", "--skip-text", f"--language={lang}", "--output-type", "pdf", "--optimize", "0", "--fast-web-view", "0", inp, outp]
             subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120)
             with open(outp, "rb") as f: return f.read()
     except subprocess.TimeoutExpired:
         st.warning("⏱️ OCR trop long → on continue sans OCR.")
     except Exception as e:
         st.warning(f"OCR ignoré ({e})")
     return src_bytes
 
-def translate_pdf_overlay(src_bytes: bytes, src: str = "fr", tgt: str = "en") -> bytes:
+def translate_pdf_overlay(
+    src_bytes: bytes,
+    src: str = "fr",
+    tgt: str = "en",
+    progress_callback: Optional[Callable[[float], None]] = None,
+) -> bytes:
     if fitz is None:
         raise RuntimeError("PyMuPDF manquant : traduction PDF indisponible.")
     doc = fitz.open(stream=src_bytes, filetype="pdf")
+    total_pages = max(1, doc.page_count)
+    processed_pages = 0
     for page in doc:
         blocks = page.get_text("blocks")
         texts = [b[4] for b in blocks if len(b) >= 5 and isinstance(b[4], str) and b[4].strip()]
-        if not texts: continue
+        if not texts:
+            processed_pages += 1
+            if progress_callback:
+                try:
+                    progress_callback(processed_pages / total_pages)
+                except Exception:
+                    pass
+            continue
         translated = translate_batch(texts, src, tgt)
         for (x0, y0, x1, y1, _txt, *_) in blocks:
             rect = fitz.Rect(x0, y0, x1, y1)
             page.add_redact_annot(rect, fill=(1, 1, 1))
         page.apply_redactions()
         def insert_text_fit(page, rect, text, fontname="Helvetica", max_size=11, min_size=6, step=0.5, align=0):
             size = max_size
             while size >= min_size:
                 used = page.insert_textbox(rect, text, fontname=fontname, fontsize=size, align=align)
                 if used >= 0: return True
                 page.add_redact_annot(rect, fill=(1, 1, 1)); page.apply_redactions()
                 size -= step
             page.insert_textbox(rect, text, fontname=fontname, fontsize=min_size, align=align); return False
         for (x0, y0, x1, y1, _txt, *_), new_text in zip(blocks, translated):
             rect = fitz.Rect(x0, y0, x1, y1)
             insert_text_fit(page, rect, new_text, max_size=11, min_size=6, step=0.5, align=0)
+        processed_pages += 1
+        if progress_callback:
+            try:
+                progress_callback(processed_pages / total_pages)
+            except Exception:
+                pass
     out = io.BytesIO(); doc.save(out); doc.close(); out.seek(0)
+    if progress_callback:
+        try:
+            progress_callback(1.0)
+        except Exception:
+            pass
     return out.read()
 
 # =================== UI ===================
 st.title("🌐 Document Translator ")
 
-src_lang = st.selectbox("Langue source", ["fr", "en", "es", "de"], index=0)
-tgt_lang = st.selectbox("Langue cible", ["en", "fr", "es", "de"], index=1)
+src_lang = st.selectbox(
+    "Langue source",
+    LANGUAGE_CODES,
+    index=LANGUAGE_CODES.index("fr") if "fr" in LANGUAGE_CODES else 0,
+    format_func=lambda code: LANGUAGE_LABELS.get(code, code),
+)
+tgt_lang = st.selectbox(
+    "Langue cible",
+    LANGUAGE_CODES,
+    index=LANGUAGE_CODES.index("en") if "en" in LANGUAGE_CODES else 0,
+    format_func=lambda code: LANGUAGE_LABELS.get(code, code),
+)
 
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
     st.caption("Astuce : le glossaire force une traduction précise. Les DNT seront laissés tels quels.")
 
 uploaded = st.file_uploader("Dépose ton fichier .docx, .pptx ou .pdf", type=["docx", "pptx", "pdf"])
 
 if uploaded:
     data = uploaded.getvalue()
     st.info(f"Fichier reçu : {uploaded.name} ({len(data)} octets)")
 
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
-                        lang_ocr = "fra" if src_lang == "fr" else src_lang
-                        ocred = ocr_pdf_with_ocrmypdf(data, lang=lang_ocr)
-                        translated = translate_pdf_overlay(ocred, src=src_lang, tgt=tgt_lang)
+                        progress_bar = st.progress(0.0)
+                        try:
+                            progress_bar.progress(0.05)
+                            lang_ocr = TESSERACT_LANG_MAP.get(src_lang, src_lang)
+                            progress_bar.progress(0.1)
+                            ocred = ocr_pdf_with_ocrmypdf(data, lang=lang_ocr)
+                            progress_bar.progress(0.35)
+
+                            def pdf_progress(step: float) -> None:
+                                base = 0.35
+                                span = 0.6
+                                progress_bar.progress(min(0.95, base + max(0.0, step) * span))
+
+                            translated = translate_pdf_overlay(
+                                ocred,
+                                src=src_lang,
+                                tgt=tgt_lang,
+                                progress_callback=pdf_progress,
+                            )
+                            progress_bar.progress(1.0)
+                        finally:
+                            progress_bar.empty()
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
-                        translated = translate_pdf_overlay(data, src=src_lang, tgt=tgt_lang)
+                        progress_bar = st.progress(0.0)
+                        try:
+                            progress_bar.progress(0.05)
+
+                            def pdf_progress(step: float) -> None:
+                                base = 0.05
+                                span = 0.9
+                                progress_bar.progress(min(0.95, base + max(0.0, step) * span))
+
+                            translated = translate_pdf_overlay(
+                                data,
+                                src=src_lang,
+                                tgt=tgt_lang,
+                                progress_callback=pdf_progress,
+                            )
+                            progress_bar.progress(1.0)
+                        finally:
+                            progress_bar.empty()
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
             st.warning("☁️ OCR désactivé en cloud. Traduction possible uniquement si le PDF est textuel.")
             if st.button("Traduire PDF (sans OCR)", key="btn_translate_pdf_cloud"):
                 with st.spinner("Traduction PDF (sans OCR)…"):
                     try:
-                        translated = translate_pdf_overlay(data, src=src_lang, tgt=tgt_lang)
+                        progress_bar = st.progress(0.0)
+                        try:
+                            progress_bar.progress(0.05)
+
+                            def pdf_progress(step: float) -> None:
+                                base = 0.05
+                                span = 0.9
+                                progress_bar.progress(min(0.95, base + max(0.0, step) * span))
+
+                            translated = translate_pdf_overlay(
+                                data,
+                                src=src_lang,
+                                tgt=tgt_lang,
+                                progress_callback=pdf_progress,
+                            )
+                            progress_bar.progress(1.0)
+                        finally:
+                            progress_bar.empty()
                         output_name = uploaded.name.replace(".pdf", f"_{tgt_lang}.pdf")
                         st.session_state.translated_bytes = translated
                         st.session_state.translated_name = output_name
                         st.session_state.translated_mime = "application/pdf"
                         save_path = save_output_file(translated, output_name)
                         st.success("✅ PDF traduit (si textuel).")
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
-                        translated = translate_pptx_preserve_styles(
-                            data, src=src_lang, tgt=tgt_lang, translate_callable=translate_batch
-                        )
+                        progress_bar = st.progress(0.0)
+                        try:
+                            progress_bar.progress(0.1)
+
+                            def ppt_progress(step: float) -> None:
+                                base = 0.1
+                                span = 0.85
+                                progress_bar.progress(min(0.95, base + max(0.0, step) * span))
+
+                            translated = translate_pptx_preserve_styles(
+                                data,
+                                src=src_lang,
+                                tgt=tgt_lang,
+                                translate_callable=translate_batch,
+                                progress_callback=ppt_progress,
+                            )
+                            progress_bar.progress(1.0)
+                        finally:
+                            progress_bar.empty()
                         output_name = uploaded.name.replace(".pptx", f"_{tgt_lang}.pptx")
                         st.session_state.translated_bytes = translated
                         st.session_state.translated_name = output_name
                         st.session_state.translated_mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                         save_path = save_output_file(translated, output_name)
                         st.success("✅ PPTX traduit. Le bouton de téléchargement est prêt ci-dessous 👇")
                         st.info(f"💾 Fichier enregistré : {save_path}")
                     except Exception as e:
                         st.error(f"Erreur PPTX: {e}")
 
 # Download button (common)
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
 st.write("- PDF : en Cloud, OCR désactivé. Les PDF scannés doivent être traités en local.")
 
EOF
)
