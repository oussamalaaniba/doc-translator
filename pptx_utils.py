import io
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

def iter_all_shapes(shapes):
    """Itère sur toutes les formes, y compris à l'intérieur des GroupShapes."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_all_shapes(shape.shapes)
        else:
            yield shape

def _collect_text_runs_from_shape(shape):
    """Retourne les runs (pptx.text.run.TextRun) d'une shape texte ou table."""
    runs = []
    # Text frames standards / placeholders
    if hasattr(shape, "has_text_frame") and shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.text and r.text.strip():
                    runs.append(r)

    # Tables (via has_table ou type TABLE)
    if getattr(shape, "has_table", False) or shape.shape_type == MSO_SHAPE_TYPE.TABLE:
        table = shape.table
        for row in table.rows:
            for cell in row.cells:
                if hasattr(cell, "text_frame") and cell.text_frame:
                    for p in cell.text_frame.paragraphs:
                        for r in p.runs:
                            if r.text and r.text.strip():
                                runs.append(r)
    return runs

def _collect_chart_runs_if_any(shape):
    """Récupère les runs accessibles dans les titres de graphiques et titres d’axes."""
    all_runs = []
    try:
        if hasattr(shape, "has_chart") and shape.has_chart:
            chart = shape.chart
            # Titre du graphique
            try:
                if chart.has_title and hasattr(chart.chart_title, "text_frame"):
                    for p in chart.chart_title.text_frame.paragraphs:
                        for r in p.runs:
                            if r.text and r.text.strip():
                                all_runs.append(r)
            except Exception:
                pass
            # Titres d'axes (selon versions de python-pptx)
            for axis_name in ("category_axis", "value_axis", "series_axis"):
                try:
                    axis = getattr(chart, axis_name, None)
                    if axis and getattr(axis, "has_title", False) and hasattr(axis.axis_title, "text_frame"):
                        for p in axis.axis_title.text_frame.paragraphs:
                            for r in p.runs:
                                if r.text and r.text.strip():
                                    all_runs.append(r)
                except Exception:
                    pass
    except Exception:
        pass
    return all_runs

def translate_pptx_preserve_styles(src_bytes, src="fr", tgt="en", translate_callable=None):
    """
    Traduit un PPTX en conservant police/taille :
    - Zones de texte, placeholders, tableaux (runs remplacés -> style conservé)
    - Objets groupés : parcourus récursivement
    - Graphiques : titres + (si exposés) titres d’axes
    - SmartArt/diagrammes : non supportés
    """
    prs = Presentation(io.BytesIO(src_bytes))
    run_refs = []

    for slide in prs.slides:
        for shape in iter_all_shapes(slide.shapes):
            run_refs.extend(_collect_text_runs_from_shape(shape))
            run_refs.extend(_collect_chart_runs_if_any(shape))

    if run_refs and translate_callable:
        batch = [r.text for r in run_refs]
        try:
            translated = translate_callable(batch, src, tgt)
        except Exception:
            translated = batch
        for r, new in zip(run_refs, translated):
            r.text = new

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio.read()
