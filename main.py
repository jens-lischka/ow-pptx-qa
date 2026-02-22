"""
OW PPTX QA Tool — Backend API
==============================
A lightweight FastAPI service that:
1. Receives slide data from the Office.js task pane
2. Runs AI-assisted brand checks using the Anthropic API
3. Returns structured issue reports

Run locally:
    uvicorn main:app --reload --port 8000

Deploy to Azure Functions or Azure App Service for production.
"""

import os
import json
from typing import Any

import anthropic
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# ── App setup ─────────────────────────────────────────────────────────────────
app = FastAPI(title="OW PPTX QA Backend", version="1.0.0")

# Allow requests from Office task pane origins
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://*.github.io", "https://localhost:3000"],
    allow_credentials=False,
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

# ── Anthropic client ──────────────────────────────────────────────────────────
# API key is read from environment variable — never hardcode it.
# Set ANTHROPIC_API_KEY in:
#   - .env file for local development
#   - GitHub Secrets for CI
#   - Azure App Settings for production
client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

# ── Request/response models ───────────────────────────────────────────────────
class FontInfo(BaseModel):
    name: str | None = None
    size: float | None = None
    color: str | None = None
    bold: bool | None = None

class Shape(BaseModel):
    name: str | None = None
    type: str | None = None
    text: str | None = None
    fonts: list[FontInfo] = []

class SlideData(BaseModel):
    index: int
    shapes: list[Shape] = []

class QARequest(BaseModel):
    slides: list[SlideData]

class Issue(BaseModel):
    type: str          # "error" | "warning" | "passed"
    rule: str
    detail: str
    slide: int | None = None
    shape: str | None = None

class QAResponse(BaseModel):
    issues: list[Issue]


# ── OW brand system prompt ────────────────────────────────────────────────────
SYSTEM_PROMPT = """You are an expert brand quality reviewer for Oliver Wyman (OW), a global management consulting firm.

Oliver Wyman brand guidelines relevant to slide content:
- Action titles: slide titles should communicate a clear finding or recommendation, not just a topic label.
  WRONG: "Market Analysis" | RIGHT: "Asian markets represent the largest growth opportunity"
- Clarity: bullet points should be concise, parallel in structure, and support the title
- Tone: professional, confident, direct — avoid jargon, hedge words ("might", "could potentially"), and passive voice
- Data labelling: all charts and graphs must have clear axis labels, units, and source references
- Consistency: terminology should be consistent across slides (e.g. don't switch between "revenue" and "sales")

You will receive structured slide data extracted from a PowerPoint presentation.
Your job is to identify content quality issues — not font/colour issues (those are handled separately).

Respond ONLY with a valid JSON array of issue objects. No explanation, no markdown, just JSON.
Each issue object must have:
  - "type": "error" or "warning"
  - "rule": short rule name (e.g. "Weak action title")
  - "detail": specific, actionable description of the issue
  - "slide": slide number (integer)
  - "shape": shape name if applicable (string or null)

Return an empty array [] if no content issues are found.
"""

# ── QA endpoint ───────────────────────────────────────────────────────────────
@app.post("/qa", response_model=QAResponse)
async def run_qa(request: QARequest):
    if not request.slides:
        raise HTTPException(status_code=400, detail="No slides provided")

    # Build a compact text representation for Claude
    slide_summary = []
    for slide in request.slides:
        shapes_text = []
        for shape in slide.shapes:
            if shape.text and shape.text.strip():
                shapes_text.append(f"  Shape '{shape.name}': {shape.text.strip()[:500]}")
        if shapes_text:
            slide_summary.append(f"Slide {slide.index}:\n" + "\n".join(shapes_text))

    if not slide_summary:
        return QAResponse(issues=[])

    user_message = "Please review the following slide content for brand quality issues:\n\n" + \
                   "\n\n".join(slide_summary)

    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2000,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_message}]
        )

        raw = response.content[0].text.strip()

        # Strip any accidental markdown fences
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]

        issues_data = json.loads(raw)
        issues = [Issue(**item) for item in issues_data]
        return QAResponse(issues=issues)

    except json.JSONDecodeError as e:
        raise HTTPException(status_code=502, detail=f"AI returned invalid JSON: {e}")
    except anthropic.APIError as e:
        raise HTTPException(status_code=502, detail=f"Anthropic API error: {e}")


# ── Health check ──────────────────────────────────────────────────────────────
@app.get("/health")
async def health():
    return {"status": "ok", "version": "1.0.0"}
