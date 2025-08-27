import base64
import json

import azure.functions as func

from generate_pptx import build_pptx, generate_pptx
from generate_pptx.models import Slide


def test_build_pptx_creates_file():
    slides = [Slide(title="Title", bullets=["one", "two"])]
    data = build_pptx(slides, "minimalist")
    assert data.startswith(b"PK")  # PPTX files are zip archives


def test_generate_route_returns_pptx():
    body = json.dumps(
        {"slides": [{"title": "Title", "bullets": ["one"]}], "style": "minimalist", "fileName": "test.pptx"}
    ).encode("utf-8")
    req = func.HttpRequest(
        method="POST",
        url="/api/generate-pptx",
        headers={},
        params={},
        route_params={},
        body=body,
    )
    resp = generate_pptx(req)
    assert resp.status_code == 200
    payload = json.loads(resp.get_body())
    assert payload["slideCount"] == 1
    base64.b64decode(payload["pptxBase64"])
