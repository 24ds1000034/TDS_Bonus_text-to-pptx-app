# Your Text, Your Style — Auto-Generate a Presentation

Turn bulk text or markdown into a fully formatted PowerPoint that inherits the style (layouts, colors, fonts, and images) from an uploaded `.pptx/.potx` template.

- **Paste text or markdown**
- **Add optional one-line guidance** (e.g., “investor pitch deck”)
- **Bring your own LLM API key** (OpenAI / Anthropic / Google Gemini)
- **Upload a PowerPoint template** to apply its look and feel
- **Download** the generated `.pptx` — *no AI image generation used*

---

## Quick start (Local)

```bash
python -m venv .venv && source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py  # runs on http://localhost:8000
```

Open your browser to `http://localhost:8000`, paste your text, add your guidance, select a provider, paste your API key, upload a `.pptx/.potx` template, and click **Generate**.

> We never store or log your API key or text. Keep Flask debug disabled in production.

---

## Deployment (One-click ideas)

- **Render**: Create a new Web Service from this repo. Set `Start Command` to `python app.py`. Use autoscaling free tier.
- **Railway / Fly.io / Deta Space**: Similar — deploy as a Python web service and expose port `8000`.
- **Self-host**: `gunicorn -w 2 -b 0.0.0.0:8000 app:app` behind Nginx.
  
No database is required. Keys are provided per request and are not stored.

---

## How it works (200–300 words)

The app uses your chosen LLM solely for **content structuring**, never for image generation. Your source text and optional one-line guidance are combined into a single prompt that instructs the LLM to output a strict JSON “slide plan” describing a sequence of slides. Each item in the plan contains a title, a short list of bullets, an optional layout hint, and optionally speaker notes if you enable them. The plan length is determined by the LLM based on content density, with a hard cap of 30 slides to preserve readability.

For rendering, the app opens your uploaded PowerPoint template (`.pptx` or `.potx`) using `python-pptx`, which preserves the theme (colors, fonts) and available layouts. For each planned slide, the app picks the best matching layout by name (e.g., “Title and Content,” “Section Header,” “Two Content”), and populates placeholders with the title and bullets. To honor the directive to reuse images from the template without generating new ones, the app scans all slides in the uploaded file to collect existing pictures and reuses them at tasteful positions, such as bottom banners or side accents on roughly every third slide or when the slide is visually sparse. If you opt in to speaker notes, notes are inserted into the slide’s notes pane. The resulting `.pptx` mirrors the look-and-feel of the original template while reflecting your text and desired tone.

---

## Security & privacy

- **API keys** are submitted with each request and are never stored or logged.
- **Request bodies** are not printed to logs. Keep Flask debug disabled in production.
- **Files** are processed in memory for transformation only and not persisted.

---

## Configuration

- Max upload: 20 MB (configurable in `app.py`).
- Providers supported:
  - **OpenAI** (default: `gpt-4o-mini`)
  - **Anthropic** (`claude-3-5-sonnet-20240620`)
  - **Google Gemini** (`gemini-1.5-pro`)
- You may specify a custom `model` name via the UI.

---

## Project structure

```
text-to-pptx-app/
├─ app.py                  # Flask web server & endpoints
├─ llm_providers.py        # Provider wrappers (OpenAI/Anthropic/Gemini)
├─ ppt_builder.py          # Template-driven slide rendering (python-pptx)
├─ requirements.txt
├─ templates/
│  └─ index.html           # Tailwind UI
├─ LICENSE                 # MIT
└─ README.md               # This file
```

---

## Notes & Limits

- The app performs best with templates that have standard “Title/Content” layouts.
- Precise layout fidelity across all vendors/themes is best-effort due to PowerPoint’s complexity.
- If a provider returns invalid JSON, the app attempts a light correction; otherwise it surfaces a friendly error.
- You can enable *speaker notes* generation via the checkbox.

---

## Roadmap / Extras (optional)

- Slide thumbnails as previews before download.
- More layout-hint coverage (e.g., timeline/process templates).
- Better image reuse heuristics (detect header/footer bands).
- Rate-limit & retry logic for unstable API responses.
```

