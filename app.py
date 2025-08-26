from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import io
import os
import json
from datetime import datetime
from ppt_builder import build_presentation
from llm_providers import plan_slides_via_llm, ProviderError

# Flask app
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20 MB upload limit
app.config['UPLOAD_EXTENSIONS'] = ['.pptx', '.potx']

# Simple "do not log sensitive content" policy: ensure debug logs are off in production
# and never print API keys or request bodies.
@app.after_request
def add_security_headers(resp):
    resp.headers['X-Content-Type-Options'] = 'nosniff'
    resp.headers['X-Frame-Options'] = 'DENY'
    resp.headers['Referrer-Policy'] = 'no-referrer'
    return resp

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    try:
        input_text = request.form.get("inputText", "").strip()
        guidance = request.form.get("guidance", "").strip()
        provider = request.form.get("provider", "openai").strip()
        model = request.form.get("model", "").strip()  # optional, allow override
        api_key = request.form.get("apiKey", "").strip()
        include_notes = request.form.get("includeNotes", "off") == "on"
        # basic validation
        if not input_text:
            return jsonify({"ok": False, "error": "Input text is required."}), 400
        if not api_key:
            return jsonify({"ok": False, "error": "API key is required for the selected provider."}), 400
        # file validation
        f = request.files.get("templateFile", None)
        if f is None or f.filename == "":
            return jsonify({"ok": False, "error": "Please upload a .pptx or .potx template/presentation."}), 400
        filename = secure_filename(f.filename)
        ext = os.path.splitext(filename)[1].lower()
        if ext not in app.config['UPLOAD_EXTENSIONS']:
            return jsonify({"ok": False, "error": "Only .pptx or .potx files are supported."}), 400
        template_bytes = f.read()

        # 1) Ask the LLM to map text -> slide plan (titles, bullets, optional notes)
        try:
            slide_plan = plan_slides_via_llm(
                provider=provider,
                model=model or None,
                api_key=api_key,
                input_text=input_text,
                guidance=guidance,
                include_notes=include_notes
            )
        except ProviderError as e:
            return jsonify({"ok": False, "error": f"LLM provider error: {e}"}), 400
        except Exception as e:
            return jsonify({"ok": False, "error": f\"Failed to get a slide plan from LLM: {e}\"}), 500

        # 2) Build PPTX from the uploaded template and the slide plan
        try:
            out_pptx = build_presentation(template_bytes, slide_plan)
        except Exception as e:
            return jsonify({"ok": False, "error": f\"Failed to build PPTX: {e}\"}), 500

        # 3) Stream the file back
        stamp = datetime.utcnow().strftime("%Y%m%d-%H%M%S")
        out_name = f"text-to-pptx-{stamp}.pptx"
        return send_file(
            io.BytesIO(out_pptx),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=out_name
        )
    except Exception as e:
        return jsonify({"ok": False, "error": f\"Unexpected error: {e}\"}), 500

# Optional: a quick JSON-only preview endpoint
@app.route("/preview", methods=["POST"])
def preview():
    try:
        input_text = request.form.get("inputText", "").strip()
        guidance = request.form.get("guidance", "").strip()
        provider = request.form.get("provider", "openai").strip()
        model = request.form.get("model", "").strip()
        api_key = request.form.get("apiKey", "").strip()
        include_notes = request.form.get("includeNotes", "off") == "on"
        if not input_text:
            return jsonify({"ok": False, "error": "Input text is required."}), 400
        if not api_key:
            return jsonify({"ok": False, "error": "API key is required for the selected provider."}), 400

        slide_plan = plan_slides_via_llm(
            provider=provider,
            model=model or None,
            api_key=api_key,
            input_text=input_text,
            guidance=guidance,
            include_notes=include_notes
        )
        return jsonify({"ok": True, "slides": slide_plan})
    except ProviderError as e:
        return jsonify({"ok": False, "error": f"LLM provider error: {e}"}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f\"Unexpected error: {e}\"}), 500

if __name__ == "__main__":
    # Never enable debug logging in production because we don't want to risk logging sensitive inputs
    app.run(host="0.0.0.0", port=8000, debug=False)
