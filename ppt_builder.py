from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
import io
import random

# Map our layout hints to likely layout names in real templates.
LAYOUT_PREFS = [
    ("title_and_content", ["Title and Content", "Content with Caption", "Title and Content (2)", "Title and Content 2"]),
    ("title_only",        ["Title Only", "Blank Title", "Title"]),
    ("section_header",    ["Section Header", "Section Title", "Title Slide"]),
    ("two_content",       ["Two Content", "Two Content and Title"]),
    ("quote",             ["Quote", "Title Only"]),
    ("comparison",        ["Comparison", "Two Content"]),
    ("timeline",          ["Title and Content", "Two Content"]),
    ("process",           ["Title and Content", "Two Content"]),
    ("overview",          ["Title and Content", "Title Only"]),
    ("summary",           ["Title and Content", "Title Only"]),
]

def _find_layout_index(prs: Presentation, layout_hint: str) -> int:
    """Pick a slide layout index from the template using our hint -> name mapping."""
    prefs = next((p[1] for p in LAYOUT_PREFS if p[0] == layout_hint), None) or ["Title and Content", "Title Only"]
    # Try preferred names
    for name in prefs:
        for i, layout in enumerate(prs.slide_layouts):
            try:
                if getattr(layout, "name", "") and name.lower() in layout.name.lower():
                    return i
            except Exception:
                pass
    # Fallback: any title layout
    for i, layout in enumerate(prs.slide_layouts):
        try:
            if getattr(layout, "name", "") and "title" in layout.name.lower():
                return i
        except Exception:
            pass
    # Last resort: first layout
    return 0

def _collect_template_images(prs: Presentation):
    """Collect (blob, width, height) of images present in the uploaded template/presentation."""
    images = []
    for slide in prs.slides:
        for shape in slide.shapes:
            try:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and hasattr(shape, "image"):
                    blob = shape.image.blob
                    images.append((blob, shape.width, shape.height))
            except Exception:
                # Skip any shapes we can’t read
                continue
    random.shuffle(images)
    return images

def _set_text(shape, text: str):
    """Safely set text into a placeholder or textbox if it has a text_frame."""
    try:
        if not hasattr(shape, "text_frame") or shape.text_frame is None:
            return
        tf = shape.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = text
        # Gentle size reduction for long titles
        if len(text) > 70:
            for r in p.runs:
                r.font.size = Pt(20)
    except Exception:
        pass

def _set_bullets(shape, bullets):
    """Write bullets into a placeholder or textbox (if it has a text_frame)."""
    try:
        if not hasattr(shape, "text_frame") or shape.text_frame is None:
            return
        tf = shape.text_frame
        tf.clear()
        for i, b in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            run = p.add_run()
            run.text = b
            p.level = 0
            if len(b) > 80:  # reduce font size if very long
                for r in p.runs:
                    r.font.size = Pt(16)
    except Exception:
        pass

def _ensure_notes(slide, text: str):
    """Add speaker notes if requested; ignore failures if template lacks notes master."""
    if not text:
        return
    try:
        notes = slide.notes_slide  # may create a notes slide if missing
        # notes can be None in some rare master configurations; guard it
        if notes and notes.notes_text_frame:
            notes.notes_text_frame.text = text
    except Exception:
        # If the template lacks a notes master or python-pptx can’t attach, just skip
        pass

def build_presentation(template_bytes: bytes, slides_plan):
    """
    Build a new PPTX from the uploaded template and an LLM-produced slide plan.
    slides_plan is a list of dicts: {title, bullets, layout_hint, notes?}
    """
    prs = Presentation(io.BytesIO(template_bytes))

    # Gather images from the template to reuse tastefully
    template_images = _collect_template_images(prs)

    for idx, slide_data in enumerate(slides_plan):
        layout_idx = _find_layout_index(prs, slide_data.get("layout_hint", "title_and_content"))
        slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])

        title_text = (slide_data.get("title") or "")[:200]
        bullets = (slide_data.get("bullets") or [])[:12]

        # Heuristic: placeholder[0] is usually title, [1] content
        title_placeholder = slide.placeholders[0] if len(slide.placeholders) > 0 else None
        content_placeholder = slide.placeholders[1] if len(slide.placeholders) > 1 else None

        if title_placeholder:
            _set_text(title_placeholder, title_text)
        else:
            # If no title placeholder, drop a textbox at top
            try:
                tb = slide.shapes.add_textbox(Inches(1), Inches(0.7), Inches(8), Inches(1))
                _set_text(tb, title_text)
            except Exception:
                pass

        if content_placeholder:
            _set_bullets(content_placeholder, bullets)
        else:
            # Add a content textbox in a safe area
            try:
                tb = slide.shapes.add_textbox(Inches(1), Inches(1.7), Inches(8), Inches(4.5))
                _set_bullets(tb, bullets)
            except Exception:
                pass

        # Speaker notes (optional)
        _ensure_notes(slide, slide_data.get("notes", ""))

        # Reuse a template image every 3rd slide or when bullets are sparse
        if template_images and ((idx % 3 == 0) or len(bullets) <= 2):
            blob, w, h = template_images[idx % len(template_images)]
            try:
                # Place as a small accent at the bottom; keep aspect ratio by constraining height
                slide.shapes.add_picture(io.BytesIO(blob), Inches(0.4), Inches(5.2), height=Inches(1.2))
            except Exception:
                # If placement fails due to layout, skip
                pass

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()
