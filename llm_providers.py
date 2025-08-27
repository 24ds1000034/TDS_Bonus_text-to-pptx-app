import os, json, re
import requests

class ProviderError(Exception):
    pass

SYSTEM_PROMPT = """You are a slide planner. Convert the provided text into a JSON slide plan.
Follow these rules:
- Respect the 'guidance' string for tone/structure if provided.
- Choose a reasonable number of slides (min 4, max 30) based on content.
- Each slide must include: title (string), bullets (array of short strings).
- If include_notes=true, add an optional notes field (string) for speaker notes.
- Use concise, scannable bullets. Avoid paragraphs.
- Do not include any images or graphics; the app will reuse template images itself.
- Output strictly valid JSON only, matching this schema:

{
  "slides": [
    {
      "title": "string",
      "bullets": ["string", "string", "..."],
      "layout_hint": "title_and_content|title_only|section_header|two_content|quote|comparison|timeline|process|overview|summary",
      "notes": "optional string"
    }
  ]
}
"""

USER_TEMPLATE = """GUIDANCE (optional): {guidance}

SOURCE TEXT:
{input_text}
"""

def _post_openai(api_key, model, system, user):
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model or "gpt-4o-mini",
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user}
        ],
        "temperature": 0.2,
        "response_format": {"type": "json_object"}
    }
    r = requests.post(url, headers=headers, json=payload, timeout=60)
    if r.status_code >= 400:
        raise ProviderError(f"OpenAI API error {r.status_code}: {r.text[:200]}")
    data = r.json()
    try:
        content = data["choices"][0]["message"]["content"]
    except Exception:
        raise ProviderError("OpenAI response missing content")
    return content

def _post_aipipe(api_key, model, system, user):
    url = "https://aipipe.org/openai/v1/chat/completions"  # OpenAI-compatible
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": model or "gpt-4o-mini",
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user}
        ],
        "temperature": 0.2,
        "response_format": {"type": "json_object"}
    }
    r = requests.post(url, headers=headers, json=payload, timeout=60)
    if r.status_code >= 400:
        raise ProviderError(f"AI-Pipe API error {r.status_code}: {r.text[:200]}")
    data = r.json()
    try:
        return data["choices"][0]["message"]["content"]
    except Exception:
        raise ProviderError("AI-Pipe response missing content")

def _post_anthropic(api_key, model, system, user):
    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }
    payload = {
        "model": model or "claude-3-5-sonnet-20240620",
        "max_tokens": 2000,
        "system": system,
        "messages": [{"role": "user", "content": user}]
    }
    r = requests.post(url, headers=headers, json=payload, timeout=60)
    if r.status_code >= 400:
        raise ProviderError(f"Anthropic API error {r.status_code}: {r.text[:200]}")
    data = r.json()
    try:
        parts = data["content"]
        text = "".join([p.get("text", "") for p in parts if p.get("type") == "text"])
    except Exception:
        raise ProviderError("Anthropic response missing text content")
    return text

def _post_gemini(api_key, model, system, user):
    model_name = model or "gemini-1.5-pro"
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
    headers = {"content-type": "application/json"}
    prompt = f"{system}\n\nUser Input:\n{user}\n\nReturn STRICT JSON only."
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.2}
    }
    r = requests.post(url, headers=headers, json=payload, timeout=60)
    if r.status_code >= 400:
        raise ProviderError(f"Gemini API error {r.status_code}: {r.text[:200]}")
    data = r.json()
    try:
        text = data["candidates"][0]["content"]["parts"][0]["text"]
    except Exception:
        raise ProviderError("Gemini response missing text content")
    return text

def _coerce_json(text):
    s = text.strip()
    if s.startswith("```"):
        s = re.sub(r"^```(?:json)?", "", s).strip()
        s = re.sub(r"```$", "", s).strip()
    if not s.startswith("{"):
        m = re.search(r"\{.*\}", s, flags=re.S)
        if m:
            s = m.group(0)
    try:
        return json.loads(s)
    except json.JSONDecodeError as e:
        raise ProviderError(f"Provider returned non-JSON output: {e}")

def _validate_api_key_like(s: str):
    bad_substrings = [" ", "http", "Bearer ", "provider.lower()", "elif "]
    if not s or any(x in s for x in bad_substrings) or len(s) < 20:
        raise ProviderError("API key looks invalid. Paste only your provider token (no quotes/Bearer/spaces).")

def plan_slides_via_llm(provider, model, api_key, input_text, guidance, include_notes):
    user = USER_TEMPLATE.format(guidance=guidance or "(none)", input_text=input_text[:15000])
    if provider.lower() == "openai":
        raw = _post_openai(api_key, model, SYSTEM_PROMPT, user)
    elif provider.lower() == "anthropic":
        raw = _post_anthropic(api_key, model, SYSTEM_PROMPT, user)
    elif provider.lower() in ("google", "gemini", "google-gemini"):
        raw = _post_gemini(api_key, model, SYSTEM_PROMPT, user)
    elif provider.lower() == "aipipe":
        raw = _post_aipipe(api_key, model, SYSTEM_PROMPT, user)
    else:
        raise ProviderError(f"Unsupported provider: {provider}")

    data = _coerce_json(raw)
    slides = data.get("slides") or []
    if not isinstance(slides, list) or len(slides) == 0:
        raise ProviderError("No slides returned by provider.")

    slides = slides[:30]
    if include_notes:
        for s in slides:
            if "notes" not in s:
                s["notes"] = ""
    for s in slides:
        if not s.get("layout_hint"):
            s["layout_hint"] = "title_and_content"
    return slides
