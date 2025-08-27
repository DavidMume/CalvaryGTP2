import base64
import io
import json
import logging
from typing import Any, Dict, List

import azure.functions as func
from pptx import Presentation
from pptx.util import Pt

logger = logging.getLogger("generate_pptx")

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)


def build_pptx(slides: List[Dict[str, Any]], style: str) -> bytes:
    """Build a PowerPoint presentation.

    Args:
        slides: List of slide definitions containing title and bullets.
        style: Styling keyword affecting font sizes.

    Returns:
        Bytes of the generated PPTX file.
    """
    prs = Presentation()
    layout = prs.slide_layouts[1]

    style_fonts = {
        "minimalist": (Pt(32), Pt(18)),
        "corporate": (Pt(36), Pt(20)),
        "colorful": (Pt(40), Pt(24)),
    }
    title_size, bullet_size = style_fonts.get(style, style_fonts["minimalist"])

    for slide in slides:
        sld = prs.slides.add_slide(layout)
        title_placeholder = sld.shapes.title
        content = sld.placeholders[1]

        title_placeholder.text = slide.get("title", "")
        title_placeholder.text_frame.paragraphs[0].font.size = title_size

        tf = content.text_frame
        tf.clear()
        bullets = slide.get("bullets", [])
        for bullet in bullets:
            p = tf.add_paragraph()
            p.text = bullet
            p.font.size = bullet_size
            p.level = 0

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


@app.function_name(name="generate_pptx")
@app.route(route="generate-pptx", methods=["POST"])
def generate_pptx(req: func.HttpRequest) -> func.HttpResponse:
    """HTTP trigger to generate a PPTX file based on JSON payload."""
    try:
        data = req.get_json()
    except ValueError:
        logger.warning("Invalid JSON payload")
        return func.HttpResponse(
            json.dumps({"error": "Invalid JSON"}),
            status_code=400,
            mimetype="application/json",
        )

    slides = data.get("slides")
    if not isinstance(slides, list) or not slides:
        logger.warning("Missing or invalid 'slides'")
        return func.HttpResponse(
            json.dumps({"error": "Invalid or missing 'slides'"}),
            status_code=400,
            mimetype="application/json",
        )

    for slide in slides:
        if not isinstance(slide, dict):
            logger.warning("Slide entry is not a dict")
            return func.HttpResponse(
                json.dumps({"error": "Invalid slide format"}),
                status_code=400,
                mimetype="application/json",
            )
        if not isinstance(slide.get("title"), str) or not isinstance(slide.get("bullets"), list):
            logger.warning("Slide missing title or bullets")
            return func.HttpResponse(
                json.dumps({"error": "Invalid slide structure"}),
                status_code=400,
                mimetype="application/json",
            )
        if any(not isinstance(b, str) for b in slide["bullets"]):
            logger.warning("Bullet is not a string")
            return func.HttpResponse(
                json.dumps({"error": "Invalid bullet format"}),
                status_code=400,
                mimetype="application/json",
            )

    style = data.get("style", "minimalist")
    file_name = data.get("fileName", "presentation.pptx")

    pptx_bytes = build_pptx(slides, style)
    pptx_b64 = base64.b64encode(pptx_bytes).decode("utf-8")

    logger.info(
        "Generated PPTX", extra={"slide_count": len(slides), "style": style}
    )

    response = {
        "status": "ok",
        "fileName": file_name,
        "slideCount": len(slides),
        "pptxBase64": pptx_b64,
    }
    return func.HttpResponse(
        json.dumps(response), status_code=200, mimetype="application/json"
    )
