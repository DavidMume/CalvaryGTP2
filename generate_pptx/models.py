from __future__ import annotations

from pydantic import BaseModel, Field, field_validator
from typing import List, Literal

MAX_SLIDES = 60
MAX_BULLETS = 5
MAX_BULLET_LENGTH = 200


class Slide(BaseModel):
    """Represents a single slide definition."""

    title: str
    bullets: List[str] = Field(default_factory=list, max_items=MAX_BULLETS)

    @field_validator("bullets", mode="before")
    def trim_bullets(cls, v: List[str]) -> List[str]:
        """Trim whitespace and restrict bullet length."""
        return [item.strip()[:MAX_BULLET_LENGTH] for item in v]


class PresentationRequest(BaseModel):
    """Request model for PPTX generation."""

    slides: List[Slide] = Field(..., min_items=1, max_items=MAX_SLIDES)
    style: Literal["minimalist", "corporate", "colorful"] = "minimalist"
    fileName: str = "presentation.pptx"
