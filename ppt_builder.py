from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import PP_PLACEHOLDER
import io
import random

# Map our layout hints to likely layout names in real templates.
LAYOUT_PREFS = [
    ("title_and_content", ["Title and Content", "Content with Caption", "Title and Content (2)", "Title and Content 2"]),
    ("title_only",        ["Title Only", "Blank Title", "Title"]),
    ("section_header",    ["Section Header", "Section Title", "Title Slide"]),
    ("two_content",       ["Two Content", "Two Content and Title", "Comparison"]),
    ("quote",             ["Quote", "Title Only"]),
    ("comparison",        ["Comparison", "Two Content"]),
    ("timeline",          ["Title and Content", "Two Content"]),
    ("process",           ["Title and Content", "Two Content"]),
    ("overview",          ["Title and Content", "Title Only"]),
    ("summary",           ["Title and Content", "Title Only"]),
]

def _find_layout_index(prs: Presentation, layout_hint: str) -> int:
    prefs = next((p[1] for p in LAYOUT_PREFS if p[0] == layout_hint), None) or ["Title and Content", "Title Only"]
    for name in prefs:
        for i, layout in enumerate(prs.slide_layouts):
            try:
                lname = getattr(layout, "name", "") or ""
                if name.lower() in lname.lower():
                    return i
            except Exception:
                pass
    for i, layout in enumerate(prs.slide_layouts):
        try:
            lname = getattr(layout, "name", "") or ""
            if "title" in lname.lower():
                return i
        except Exception:
            pass
    return 0

def _collect_template_images(prs: Presentation):
    images = []
    for slide in prs.slides:
        for shape in slide.shapes:
            try:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and hasattr(shape, "image"):
                    images.append((shape.image.blob, shape.width, shape.height))
            except Exception:
                continue
    random.shuffle(images)
    return images

def _purge_all_existing_slides(prs: Presentation):
    sldIdLst = prs.slides._sldIdLst
    slide_ids = list(sldIdLst)
    for sldId in slide_ids:
        rId = sldId.rId
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)

def _first_placeholder(slide, *types):
    for shp in slide.placeholders:
        try:
            if shp.placeholder_format and shp.placeholder_format.type in types:
                return shp
        except Exception:
            pass
    return None

def _first_non_title_text_placeholder(slide):
    for shp in slide.placeholders:
        try:
            t = shp.placeholder_format.type
            if t not in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.SUBTITLE):
                if hasattr(shp, "text_frame") and shp.text_frame is not None:
                    return shp
        except Exception:
            pass
    return None

def _set_text(shape, text: str):
    try:
        if not hasattr(shape, "text_frame") or shape.text_frame is None:
            return
        tf = shape.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = text or ""
        if len(text or "") > 70:
            for r in p.runs:
                r.font.size = Pt(20)
    except Exception:
        pass

def _set_bullets(shape, bullets):
    try:
        if not hasattr(shape, "text_frame") or shape.text_frame is None:
            return
        tf = shape.text_frame
        tf.clear()
        bullets = bullets or []
        for i, b in enumerate(bullets[:12]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            run = p.add_run()
            run.text = b
            p.level = 0
            if len(b) > 80:
                for r in p.runs:
                    r.font.size = Pt(16)
    except Exception:
        pass

def _ensure_notes(slide, text: str):
    if not text:
        return
    try:
        notes = slide.notes_slide
        if notes and notes.notes_text_frame:
            notes.notes_text_frame.text = text
    except Exception:
        pass

def build_presentation(template_bytes: bytes, slides_plan):
    prs = Presentation(io.BytesIO(template_bytes))

    # 1) Collect images then purge slides so template content doesn't leak in.
    template_images = _collect_template_images(prs)
    _purge_all_existing_slides(prs)

    # 2) Build slides from plan
    for idx, slide_data in enumerate(slides_plan):
        layout_idx = _find_layout_index(prs, slide_data.get("layout_hint", "title_and_content"))
        slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])

        title_text = (slide_data.get("title") or "")[:200]
        bullets = (slide_data.get("bullets") or [])[:12]

        title_ph = _first_placeholder(slide, PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.SUBTITLE)
        body_ph  = _first_placeholder(slide, PP_PLACEHOLDER.BODY) or _first_non_title_text_placeholder(slide)

        if title_ph:
            _set_text(title_ph, title_text)
        else:
            try:
                tb = slide.shapes.add_textbox(Inches(1), Inches(0.7), Inches(8), Inches(1))
                _set_text(tb, title_text)
            except Exception:
                pass

        if body_ph:
            _set_bullets(body_ph, bullets)
        else:
            try:
                tb = slide.shapes.add_textbox(Inches(1), Inches(1.7), Inches(8), Inches(4.5))
                _set_bullets(tb, bullets)
            except Exception:
                pass

        _ensure_notes(slide, slide_data.get("notes", ""))

        if template_images and ((idx % 3 == 0) or len(bullets) <= 2):
            blob, w, h = template_images[idx % len(template_images)]
            try:
                slide.shapes.add_picture(io.BytesIO(blob), Inches(0.4), Inches(5.2), height=Inches(1.2))
            except Exception:
                pass

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()
