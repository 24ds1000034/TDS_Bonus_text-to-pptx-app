from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
import io
import random

# Simple mapping from layout_hint -> best-effort pick from template
LAYOUT_PREFS = [
    ("title_and_content", ["Title and Content", "Content with Caption", "Title and Content (2)", "Title and Content 2"]),
    ("title_only",        ["Title Only", "Blank Title", "Title"]),
    ("section_header",    ["Section Header", "Section Title", "Title Slide"]),
    ("two_content",       ["Two Content", "Two Content and Title"]),
    ("quote",             ["Quote", "Title Only"]),  # fallback
    ("comparison",        ["Comparison", "Two Content"]),
    ("timeline",          ["Title and Content", "Two Content"]),
    ("process",           ["Title and Content", "Two Content"]),
    ("overview",          ["Title and Content", "Title Only"]),
    ("summary",           ["Title and Content", "Title Only"]),
]

def _find_layout_index(prs, layout_hint):
    prefs = next((p[1] for p in LAYOUT_PREFS if p[0] == layout_hint), None)
    if prefs is None:
        prefs = ["Title and Content", "Title Only"]
    # try to match by name
    for name in prefs:
        for i, l in enumerate(prs.slide_layouts):
            try:
                if l.name and name.lower() in l.name.lower():
                    return i
            except Exception:
                pass
    # fallback to a generic content layout
    for i, l in enumerate(prs.slide_layouts):
        try:
            if l.name and "title" in l.name.lower():
                return i
        except Exception:
            pass
    return 0  # last resort

def _collect_template_images(prs):
    images = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    # extract the image bytes
                    image_blob = shape.image.blob
                    # size similar to original
                    width = shape.width
                    height = shape.height
                    images.append((image_blob, width, height))
                except Exception:
                    pass
    return images

def _add_text_to_placeholder(shape, text):
    try:
        tf = shape.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = text
        p.level = 0
        # slightly reduce font size if very long
        for r in p.runs:
            if len(text) > 70:
                r.font.size = Pt(20)
    except Exception:
        # ignore if not a text placeholder
        pass

def _add_bullets_to_placeholder(shape, bullets):
    try:
        tf = shape.text_frame
        tf.clear()
        # title often occupies placeholder 0; bullets go to the body placeholder
        for i, b in enumerate(bullets):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            run = p.add_run()
            run.text = b
            p.level = 0
            # If bullet length is big, adjust size
            for r in p.runs:
                if len(b) > 80:
                    r.font.size = Pt(16)
    except Exception:
        pass

def build_presentation(template_bytes, slides_plan):
    # Start from the uploaded template so we inherit theme, colors, fonts, layouts
    prs = Presentation(io.BytesIO(template_bytes))

    # Try to collect images from template to reuse
    template_images = _collect_template_images(prs)
    # Put a gentle cap so we don't overuse images
    random.shuffle(template_images)

    for idx, slide_data in enumerate(slides_plan):
        layout_idx = _find_layout_index(prs, slide_data.get("layout_hint", "title_and_content"))
        slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])

        # Heuristic: assume placeholder[0] is title, placeholder[1] is content (common in many templates)
        title_text = slide_data.get("title", "")[:200]
        bullets = slide_data.get("bullets", [])[:12]

        # Fill placeholders where available
        if len(slide.placeholders) > 0:
            _add_text_to_placeholder(slide.placeholders[0], title_text)
        # find next best placeholder for bullets
        target_for_bullets = None
        if len(slide.placeholders) > 1:
            target_for_bullets = slide.placeholders[1]
        else:
            # fallback: add a new textbox
            x, y, w, h = Inches(1), Inches(1.7), Inches(8), Inches(4.5)
            target_for_bullets = slide.shapes.add_textbox(x, y, w, h)
        _add_bullets_to_placeholder(target_for_bullets, bullets)

        # Optional: add speaker notes
        notes_text = slide_data.get("notes", "")
        if notes_text:
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = notes_text

        # Reuse a template image approximately every 3rd slide or when bullets are light
        if template_images and ((idx % 3 == 0) or len(bullets) <= 2):
            image_blob, w, h = template_images[idx % len(template_images)]
            # Place at bottom or side depending on layout width; keep a modest size
            try:
                pic = slide.shapes.add_picture(io.BytesIO(image_blob), Inches(0.4), Inches(5.2), height=Inches(1.2))
            except Exception:
                # Ignore if placement fails
                pass

    # Save final file to bytes
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()
