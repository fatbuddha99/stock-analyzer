from __future__ import annotations

import io
import os
import tempfile
from contextlib import redirect_stdout
from pathlib import Path

from flask import Flask, jsonify, render_template_string, request
from werkzeug.utils import secure_filename

import stock_analyzer


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

INDEX_HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Stock Analyzer</title>
  <style>
    :root {
      --bg: #0b1220;
      --panel: #131c2f;
      --panel-2: #1a2742;
      --text: #e6edf8;
      --muted: #93a4bf;
      --accent: #56d7a4;
      --danger: #ff7a7a;
      --border: #2a3958;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Segoe UI", Tahoma, sans-serif;
      background: radial-gradient(circle at top, #1b2a4a 0%, var(--bg) 52%);
      color: var(--text);
      min-height: 100vh;
    }
    .wrap {
      max-width: 920px;
      margin: 0 auto;
      padding: 40px 20px 64px;
    }
    .hero {
      margin-bottom: 28px;
    }
    h1 {
      margin: 0 0 10px;
      font-size: 2.4rem;
      line-height: 1.05;
    }
    .sub {
      color: var(--muted);
      max-width: 720px;
      line-height: 1.5;
    }
    .card {
      background: linear-gradient(180deg, rgba(255,255,255,0.04), rgba(255,255,255,0.02));
      border: 1px solid var(--border);
      border-radius: 18px;
      padding: 22px;
      backdrop-filter: blur(8px);
      box-shadow: 0 18px 45px rgba(0, 0, 0, 0.28);
    }
    form {
      display: grid;
      gap: 16px;
    }
    .grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
      gap: 16px;
    }
    label {
      display: grid;
      gap: 8px;
      color: var(--muted);
      font-size: 0.95rem;
    }
    input, select, button {
      width: 100%;
      border-radius: 12px;
      border: 1px solid var(--border);
      background: var(--panel-2);
      color: var(--text);
      padding: 13px 14px;
      font: inherit;
    }
    button {
      background: linear-gradient(135deg, #2fbf8a, var(--accent));
      color: #072116;
      border: 0;
      font-weight: 700;
      cursor: pointer;
    }
    button:hover { filter: brightness(1.03); }
    .hint {
      color: var(--muted);
      font-size: 0.9rem;
      margin-top: -4px;
    }
    .error {
      margin-top: 18px;
      color: var(--danger);
      white-space: pre-wrap;
    }
    pre {
      margin: 18px 0 0;
      padding: 18px;
      border-radius: 14px;
      background: #08101d;
      color: #dfe8f6;
      border: 1px solid var(--border);
      overflow-x: auto;
      white-space: pre-wrap;
      line-height: 1.4;
    }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="hero">
      <h1>Stock Analyzer</h1>
      <div class="sub">
        Run the Minervini + Kell technical analyzer on demand. This Render wrapper uses
        <code>yfinance</code> by default, supports an optional chart upload, and returns the
        same terminal brief you see locally.
      </div>
    </div>
    <div class="card">
      <form method="post" enctype="multipart/form-data">
        <div class="grid">
          <label>
            Ticker
            <input type="text" name="ticker" placeholder="AAPL" required value="{{ ticker }}">
          </label>
          <label>
            Period
            <select name="period">
              {% for item in periods %}
              <option value="{{ item }}" {% if item == period %}selected{% endif %}>{{ item }}</option>
              {% endfor %}
            </select>
          </label>
          <label>
            Save DOCX
            <select name="docx">
              <option value="0" {% if not save_docx %}selected{% endif %}>No</option>
              <option value="1" {% if save_docx %}selected{% endif %}>Yes</option>
            </select>
          </label>
        </div>
        <label>
          Chart Image (optional)
          <input type="file" name="chart" accept=".png,.jpg,.jpeg,.webp">
        </label>
        <div class="hint">
          Vision analysis requires the <code>ANTHROPIC_API_KEY</code> environment variable in Render.
        </div>
        <button type="submit">Analyze</button>
      </form>
      {% if error %}
      <div class="error">{{ error }}</div>
      {% endif %}
      {% if output %}
      <pre>{{ output }}</pre>
      {% endif %}
      {% if download_url %}
      <div class="hint" style="margin-top: 14px;">
        DOCX created: <a href="{{ download_url }}" style="color: #8ed9ff;">{{ download_url }}</a>
      </div>
      {% endif %}
    </div>
  </div>
</body>
</html>
"""


def run_analysis(ticker: str, period: str, chart_path: str | None, save_docx: bool):
    ticker = ticker.upper().strip()
    output_dir = Path("outputs")
    output_dir.mkdir(exist_ok=True)
    output_path = output_dir / f"{ticker}_analysis.docx" if save_docx else None

    buffer = io.StringIO()
    with redirect_stdout(buffer):
        tag = stock_analyzer.run(
            ticker,
            output_path=str(output_path) if output_path else None,
            period=period,
            source="yfinance",
            chart_path=chart_path,
            api_key=os.environ.get("ANTHROPIC_API_KEY"),
            save_docx=save_docx,
        )
    return {
        "ticker": ticker,
        "tag": tag,
        "output": buffer.getvalue(),
        "docx_path": str(output_path) if output_path and output_path.exists() else None,
    }


@app.get("/health")
def health():
    return jsonify({"ok": True})


@app.route("/", methods=["GET", "POST"])
def index():
    output = ""
    error = ""
    download_url = ""
    ticker = ""
    period = "3y"
    save_docx = False
    temp_chart_path = None

    if request.method == "POST":
        ticker = request.form.get("ticker", "").strip().upper()
        period = request.form.get("period", "3y")
        save_docx = request.form.get("docx", "0") == "1"
        chart = request.files.get("chart")

        if chart and chart.filename:
            suffix = Path(secure_filename(chart.filename)).suffix or ".png"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                chart.save(tmp)
                temp_chart_path = tmp.name

        try:
            result = run_analysis(ticker, period, temp_chart_path, save_docx)
            output = result["output"]
            if result["docx_path"]:
                download_url = f"/download/{Path(result['docx_path']).name}"
        except Exception as exc:
            error = f"{type(exc).__name__}: {exc}"
        finally:
            if temp_chart_path and os.path.exists(temp_chart_path):
                os.unlink(temp_chart_path)

    return render_template_string(
        INDEX_HTML,
        output=output,
        error=error,
        ticker=ticker,
        period=period,
        save_docx=save_docx,
        download_url=download_url,
        periods=["1y", "2y", "3y", "5y", "10y", "max"],
    )


@app.post("/api/analyze")
def analyze():
    ticker = request.form.get("ticker", request.args.get("ticker", "")).strip().upper()
    period = request.form.get("period", request.args.get("period", "3y"))
    save_docx = request.form.get("docx", request.args.get("docx", "0")) in {"1", "true", "True"}
    chart = request.files.get("chart")
    temp_chart_path = None

    if not ticker:
        return jsonify({"ok": False, "error": "ticker is required"}), 400

    if chart and chart.filename:
        suffix = Path(secure_filename(chart.filename)).suffix or ".png"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            chart.save(tmp)
            temp_chart_path = tmp.name

    try:
        result = run_analysis(ticker, period, temp_chart_path, save_docx)
        payload = {
            "ok": True,
            "ticker": result["ticker"],
            "tag": result["tag"],
            "output": result["output"],
        }
        if result["docx_path"]:
            payload["docx_download"] = f"/download/{Path(result['docx_path']).name}"
        return jsonify(payload)
    except Exception as exc:
        return jsonify({"ok": False, "error": f"{type(exc).__name__}: {exc}"}), 500
    finally:
        if temp_chart_path and os.path.exists(temp_chart_path):
            os.unlink(temp_chart_path)


@app.get("/download/<filename>")
def download(filename: str):
    from flask import send_from_directory

    return send_from_directory("outputs", filename, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
