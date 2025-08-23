import io
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

def iter_all_shapes(shapes):
    """Itère sur toutes les formes, y compris l'intérieur des GroupShapes."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_all_shapes(shape.shapes)
        else:
            yield shape

def _paragraph_text(paragraph):
    """Reconstruit le texte d'un paragraphe en concaténant les runs (préserve les espaces)."""
    if paragraph.runs:
        return "".join(r.text for r in paragraph.runs)
    return paragraph.text or ""

def _replace_paragraph_text_keep_font(paragraph, new_text):
    """Remplace TOUT le paragraphe par un seul run en conservant police/taille du 1er run."""
    first_run = paragraph.runs[0] if paragraph.runs else None
    # supprimer tous les runs existants
    for _ in range(len(paragraph.runs)):
        r = paragraph.runs[0]
        r._element.getparent().remove(r._element)
    # écrire un seul run
    new_run = paragraph.add_run()
    new_run.text = new_text
    # recopier police/taille
    if first_run:
        try:
            if first_run.font.name:
                new_run.font.name = first_run.font.name
            if first_run.font.size:
                new_run.font.size = first_run.font.size
        except Exception:
            pass  # si police de thème, on ignore

def _translate_text_frame_by_paragraph(text_frame, translate_callable, src, tgt):
    """Traduit chaque PARAGRAPHE (pas run) pour éviter les pertes d'espaces et le mot-à-mot."""
    for p in list(text_frame.paragraphs):
        original = _paragraph_text(p)
        if not original or not original.strip():
            continue
        translated = translate_callable([original], src, tgt)[0] if translate_callable else original
        _replace_paragraph_text_keep_font(p, translated)

def _translate_table_by_paragraph(table, translate_callable, src, tgt):
    for row in table.rows:
        for cell in row.cells:
            if hasattr(cell, "text_frame") and cell.text_frame:
                _translate_text_frame_by_paragraph(cell.text_frame, translate_callable, src, tgt)

def _translate_chart_titles_if_any(shape, translate_callable, src, tgt):
    """Traduit titres de graphiques et titres d’axes si exposés par python-pptx."""
    try:
        if hasattr(shape, "has_chart") and shape.has_chart:
            chart = shape.chart
            # Titre du chart
            if chart.has_title and hasattr(chart.chart_title, "text_frame"):
                _translate_text_frame_by_paragraph(chart.chart_title.text_frame, translate_callable, src, tgt)
            # Titres d’axes
            for axis_name in ("category_axis", "value_axis", "series_axis"):
                axis = getattr(chart, axis_name, None)
                if axis and getattr(axis, "has_title", False) and hasattr(axis.axis_title, "text_frame"):
                    _translate_text_frame_by_paragraph(axis.axis_title.text_frame, translate_callable, src, tgt)
    except Exception:
        pass

def translate_pptx_preserve_styles(src_bytes, src="fr", tgt="en", translate_callable=None):
    """
    Traduction PAR PARAGRAPHE (évite perte d'espaces et mot-à-mot).
    - Zones de texte / placeholders
    - Objets groupés (récursif)
    - Tableaux
    - Titres de graphiques (+ axes si exposés)
    - SmartArt/images avec texte: non supportés
    """
    prs = Presentation(io.BytesIO(src_bytes))

    for slide in prs.slides:
        for shape in iter_all_shapes(slide.shapes):
            if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                _translate_text_frame_by_paragraph(shape.text_frame, translate_callable, src, tgt)

            if getattr(shape, "has_table", False) or shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                _translate_table_by_paragraph(shape.table, translate_callable, src, tgt)

            _translate_chart_titles_if_any(shape, translate_callable, src, tgt)

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio.read()

