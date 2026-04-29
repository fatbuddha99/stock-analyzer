from __future__ import annotations

import io
import os
import tempfile
from contextlib import redirect_stdout
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template_string, request, send_from_directory
from werkzeug.utils import secure_filename

import stock_analyzer


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

if getattr(stock_analyzer, "HAS_YFINANCE", False):
    cache_dir = BASE_DIR / ".yfinance-cache"
    cache_dir.mkdir(exist_ok=True)
    try:
        stock_analyzer.yf.set_tz_cache_location(str(cache_dir))
    except Exception:
        pass

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

INDEX_HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>The Oracle</title>
  <style>
    :root {
      --bg: #020403;
      --panel: #07110b;
      --line: rgba(116, 255, 145, 0.18);
      --text: rgba(231, 238, 231, 0.84);
      --muted: rgba(169, 208, 173, 0.72);
      --green: #79ff8d;
      --orange: #ffb24a;
      --red: #ff7c9b;
      --blue: #8df5ff;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background:
        radial-gradient(circle at top, rgba(68, 255, 120, 0.08), transparent 24%),
        linear-gradient(180deg, #010302 0%, #030705 100%);
      color: var(--text);
      font-family: "Courier New", Courier, monospace;
    }
    body::before {
      content: "";
      position: fixed;
      inset: 0;
      pointer-events: none;
      opacity: 0.24;
      background:
        linear-gradient(rgba(121, 255, 141, 0.035) 1px, transparent 1px),
        linear-gradient(90deg, rgba(121, 255, 141, 0.02) 1px, transparent 1px);
      background-size: 100% 4px, 4px 100%;
    }
    .shell {
      position: relative;
      max-width: 1320px;
      margin: 0 auto;
      padding: 18px;
    }
    .topbar, .panel, .output, .chart-panel, .card, .nav-row {
      background: rgba(7, 17, 11, 0.92);
      border: 1px solid var(--line);
      border-radius: 18px;
      box-shadow: 0 0 0 1px rgba(24, 79, 36, 0.16) inset, 0 18px 50px rgba(0, 0, 0, 0.34);
    }
    .topbar {
      padding: 16px 18px;
      margin-bottom: 14px;
    }
    .title {
      margin: 0;
      font-family: Impact, Haettenschweiler, "Arial Narrow Bold", sans-serif;
      font-size: 1.9rem;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: #010402;
      -webkit-text-stroke: 1.4px var(--green);
      text-shadow: 0 0 12px rgba(121, 255, 141, 0.22);
    }
    .subtitle {
      margin-top: 8px;
      color: var(--muted);
      font-size: 0.88rem;
    }
    .nav-row {
      display: flex;
      gap: 8px;
      padding: 8px;
      margin-bottom: 14px;
      overflow-x: auto;
    }
    .nav-row::-webkit-scrollbar { display: none; }
    .btn {
      flex: 0 0 auto;
      border: 1px solid rgba(240, 244, 240, 0.16);
      border-radius: 999px;
      background: rgba(238, 244, 238, 0.08);
      color: rgba(245, 248, 245, 0.8);
      padding: 8px 12px;
      font-size: 0.72rem;
      font-family: Verdana, Geneva, sans-serif;
      text-transform: uppercase;
      letter-spacing: 0.1em;
      text-decoration: none;
      cursor: pointer;
    }
    .btn.primary {
      color: var(--green);
      border-color: rgba(121, 255, 141, 0.28);
      background: rgba(38, 87, 48, 0.25);
      font-family: "Trebuchet MS", Helvetica, sans-serif;
      font-weight: 700;
    }
    .layout {
      display: grid;
      grid-template-columns: 320px minmax(0, 1fr);
      gap: 14px;
    }
    .panel {
      padding: 14px;
      align-self: start;
    }
    .field {
      display: grid;
      gap: 8px;
      margin-bottom: 12px;
    }
    .field label {
      color: var(--muted);
      font-size: 0.76rem;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      font-family: Verdana, Geneva, sans-serif;
    }
    .text-input, .file-input {
      width: 100%;
      border: 1px solid rgba(240, 244, 240, 0.16);
      border-radius: 12px;
      background: rgba(238, 244, 238, 0.05);
      color: var(--text);
      padding: 10px 12px;
      font: inherit;
    }
    .button-row {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }
    .helper {
      color: var(--muted);
      font-size: 0.82rem;
      line-height: 1.5;
      margin-top: 8px;
    }
    .workspace, .sections {
      display: grid;
      gap: 14px;
    }
    .hero {
      display: grid;
      gap: 12px;
      padding: 16px;
      background: linear-gradient(180deg, rgba(8, 20, 11, 0.95), rgba(5, 12, 7, 0.95));
    }
    .verdict {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 6px 10px;
      border: 1px solid var(--line);
      border-radius: 999px;
      width: fit-content;
      font-size: 0.76rem;
      background: rgba(12, 28, 15, 0.75);
      color: var(--orange);
    }
    .dot {
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: var(--orange);
      box-shadow: 0 0 10px rgba(121, 255, 141, 0.55);
    }
    .ticker-row {
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 10px;
    }
    .ticker {
      font-size: 2rem;
      font-weight: 700;
      line-height: 1;
      color: #f1fff2;
      font-family: Georgia, "Times New Roman", serif;
      letter-spacing: 0.03em;
    }
    .meta {
      color: var(--muted);
      font-size: 0.8rem;
      text-align: right;
    }
    .chart-panel {
      overflow: hidden;
      min-height: 420px;
    }
    .chart-wrap {
      height: 420px;
      background: #050b06;
    }
    .output {
      padding: 0;
      overflow: hidden;
    }
    .output-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      padding: 12px 14px;
      border-bottom: 1px solid var(--line);
      color: var(--muted);
      font-size: 0.78rem;
      font-family: Verdana, Geneva, sans-serif;
      text-transform: uppercase;
      letter-spacing: 0.08em;
    }
    .status { color: var(--orange); }
    .terminal {
      margin: 0;
      padding: 16px;
      min-height: 360px;
      white-space: pre-wrap;
      overflow-x: auto;
      line-height: 1.55;
      color: var(--text);
    }
    .card {
      padding: 16px;
    }
    .section-title h2 {
      margin: 0 0 10px;
      font-size: 0.98rem;
      color: var(--green);
      font-family: "Trebuchet MS", Helvetica, sans-serif;
      letter-spacing: 0.03em;
    }
    .error { color: var(--red); }
    .download-link { color: var(--blue); }
    @media (max-width: 980px) {
      .layout { grid-template-columns: 1fr; }
      .chart-wrap { height: 360px; }
    }
    @media (max-width: 640px) {
      .shell { padding: 12px 10px 32px; }
      .ticker { font-size: 1.6rem; }
      .ticker-row { flex-direction: column; align-items: flex-start; }
      .meta { text-align: left; }
    }
  </style>
</head>
<body>
  <div class="shell">
    <div class="topbar">
      <h1 class="title">The Oracle</h1>
      <div class="subtitle">A Deep Dive For Context And Nuances.</div>
    </div>

    <div class="nav-row">
      <a class="btn" href="#insight">Insight</a>
      <a class="btn" href="#stage">Stage</a>
      <a class="btn" href="#volatility">Volatility</a>
      <a class="btn" href="#earnings">Earnings</a>
      <a class="btn" href="#action">Action</a>
    </div>

    <div class="layout">
      <form class="panel" method="post" enctype="multipart/form-data">
        <div class="field">
          <label for="ticker">Ticker</label>
          <input class="text-input" id="ticker" name="ticker" placeholder="IBM" value="{{ ticker }}">
        </div>
        <div class="field">
          <label for="chart">Annotated Chart (Optional)</label>
          <input class="file-input" id="chart" type="file" name="chart" accept=".png,.jpg,.jpeg,.webp">
        </div>
        <div class="button-row">
          <button class="btn primary" type="submit" name="action" value="consult">Consult</button>
          <button class="btn" type="submit" name="action" value="export">Export</button>
        </div>
        <div class="helper">
          Render now mirrors the localhost terminal flow and uses Yahoo-backed history instead of Alpha Vantage.
        </div>
      </form>

      <div class="workspace">
        <div class="hero">
          <div class="verdict"><span class="dot"></span> {% if show_live %}Oracle Consult{% else %}Stand By{% endif %}</div>
          <div class="ticker-row">
            <div class="ticker">{% if show_live %}{{ ticker }}{% else %}The Oracle Is Self Fufilling{% endif %}</div>
            <div class="meta">{% if show_live %}{{ consult_time }}{% else %}Waiting for a ticker consult{% endif %}</div>
          </div>
          <p>{% if show_live %}The Oracle has generated a live consult for {{ ticker }}. The chart, terminal output, and five-section read below are now the active view.{% else %}Nothing loads until you ask for a read.{% endif %}</p>
        </div>

        <div class="chart-panel">
          {% if show_live %}
          <div class="chart-wrap">
            <iframe
              title="TradingView Chart"
              src="https://s.tradingview.com/widgetembed/?symbol=NASDAQ%3A{{ chart_symbol }}&interval=D&hidesidetoolbar=1&symboledit=1&saveimage=0&toolbarbg=050b06&studies=[]&theme=dark&style=1&timezone=Etc%2FUTC&withdateranges=1&hideideas=1&hidevolume=0&calendar=0&details=0&hotlist=0&news=0&watchlist=0&locale=en"
              style="width:100%;height:100%;border:0;"
              allowtransparency="true"
              scrolling="no"></iframe>
          </div>
          {% else %}
          <pre class="terminal">The Oracle is waiting for a ticker.</pre>
          {% endif %}
        </div>

        <div class="output">
          <div class="output-head">
            <span>Terminal Output</span>
            <span class="status">{% if show_live %}Consult Complete{% else %}Stand By{% endif %}</span>
          </div>
          <pre class="terminal{% if error %} error{% endif %}">{{ output if output else 'The Oracle is waiting for a consult.' }}</pre>
          {% if download_url %}
          <p style="margin:0;padding:0 16px 16px;">DOCX created: <a class="download-link" href="{{ download_url }}">{{ download_url }}</a></p>
          {% endif %}
        </div>
      </div>
    </div>

    {% if show_live %}
    <div class="sections" style="margin-top:14px;">
      <div class="card" id="insight">
        <div class="section-title"><h2>Insight</h2></div>
        <p>{{ insight_text }}</p>
      </div>
      <div class="card" id="stage">
        <div class="section-title"><h2>Stage Analysis</h2></div>
        <p>{{ stage_text }}</p>
      </div>
      <div class="card" id="volatility">
        <div class="section-title"><h2>Volatility Analysis</h2></div>
        <p>{{ volatility_text }}</p>
      </div>
      <div class="card" id="earnings">
        <div class="section-title"><h2>Earnings Context</h2></div>
        <p>{{ earnings_text }}</p>
      </div>
      <div class="card" id="action">
        <div class="section-title"><h2>Action Plan</h2></div>
        <p>{{ action_text }}</p>
      </div>
    </div>
    {% elif error %}
    <div class="card" style="margin-top:14px;">
      <div class="section-title"><h2>Error</h2></div>
      <p class="error">{{ error }}</p>
    </div>
    {% else %}
    <div class="card" id="insight" style="margin-top:14px;">
      <div class="section-title"><h2>Insight</h2></div>
      <p>The Oracle waits for a ticker, then builds a live consult from chart structure, trend state, volatility character, and the script's own action plan. Nothing loads until you ask for a read.</p>
    </div>
    {% endif %}
  </div>
</body>
</html>
"""


def build_render_sections(
    ticker: str,
    tag: str,
    structure: str,
    stage_info: dict,
    vcp: dict,
    kell: dict,
    fib: dict | None,
    vah: str,
    action: str,
    bull: list[str],
    bear: list[str],
    ticker_obj,
) -> dict[str, str]:
    stage_name = stage_info.get("stage", "Unknown")
    vcp_grade = vcp.get("grade", "Unknown")
    retrace = fib.get("retrace_pct", 0) if fib else 0
    kell_pattern = kell.get("pattern", "Neutral / Consolidation")

    insight_text = (
        f"{ticker} currently reads as {tag.lower()}, with {len(bull)} bullish signals versus "
        f"{len(bear)} bearish signals in the script output. The immediate takeaway is that The Oracle "
        f"sees the tape as {structure.lower()}, not as a fresh momentum continuation."
    )
    stage_text = (
        f"The stage model is {stage_name}, and the underlying swing structure is {structure.lower()}. "
        "That means the read is being driven more by broad trend condition than by a clean breakout posture."
    )
    volatility_text = (
        f"Volatility quality is {vcp_grade}, while the Kell-pattern read is {kell_pattern.lower()}. "
        f"The latest Fibonacci retracement is near {retrace:.1f}%, and the weekly value-area read is: {vah}."
    )

    earnings_text = (
        "This Render version now mirrors the localhost terminal and does not use Alpha Vantage. "
        "The live read is primarily technical, with any catalyst context coming from Yahoo-backed news the script can fetch."
    )
    if ticker_obj:
        try:
            news = stock_analyzer.fetch_news(ticker_obj, 21)
            if news:
                latest = news[0]
                earnings_text = (
                    f"Recent context from Yahoo-backed news is available. The latest headline was on "
                    f"{latest['date'].strftime('%b %d, %Y')} from {latest['publisher']}: {latest['title']}. "
                    "This section is now driven by the same lightweight Yahoo path as localhost."
                )
        except Exception:
            pass

    return {
        "insight_text": insight_text,
        "stage_text": stage_text,
        "volatility_text": volatility_text,
        "earnings_text": earnings_text,
        "action_text": action,
    }


def run_analysis(ticker: str, period: str, chart_path: str | None, save_docx: bool):
    ticker = ticker.upper().strip()
    output_path = OUTPUT_DIR / f"{ticker}_analysis.docx" if save_docx else None

    df, ticker_obj = stock_analyzer.fetch_data(ticker, period, source="yfinance")
    df = stock_analyzer.add_indicators(df)
    df = stock_analyzer.classify_volume_colors(df)

    sw_hi, sw_lo = stock_analyzer.find_swings(df)
    structure, prior_hi, prior_lo = stock_analyzer.determine_structure(df, sw_hi, sw_lo)
    reclaim = stock_analyzer.determine_reclaim_levels(df, sw_hi)
    stage_info = stock_analyzer.determine_stage(df)
    vcp = stock_analyzer.vcp_scorecard(df)
    kell = stock_analyzer.kell_cycle(df)
    fib = stock_analyzer.fib_retracement(df)
    profile = stock_analyzer.weekly_profile(df)
    vah = stock_analyzer.vah_ladder_analysis(profile)
    news = stock_analyzer.fetch_news(ticker_obj, 45) if ticker_obj else []
    ep = stock_analyzer.episodic_pivot_check(df, news)
    bull, bear, action, tag = stock_analyzer.classify_signals(
        df, stage_info, vcp, kell, fib, vah, structure, prior_hi, prior_lo, reclaim, ep
    )

    vision_md = None
    if chart_path:
        lines = [
            f"Verdict tag: {tag}",
            f"Bullish: {' | '.join(bull) if bull else 'none'}",
            f"Bearish: {' | '.join(bear) if bear else 'none'}",
            f"Reclaim Tiers: T1 ${reclaim['tier1']['price']:.2f}, "
            f"T2 ${reclaim['tier2']['price']:.2f}, T3 ${reclaim['tier3']['price']:.2f}",
        ]
        vision_md = stock_analyzer.analyze_chart_with_vision(
            chart_path, ticker, "\n".join(lines), api_key=os.environ.get("ANTHROPIC_API_KEY")
        )

    buffer = io.StringIO()
    with redirect_stdout(buffer):
        stock_analyzer.print_brief(ticker, df, bull, bear, action, tag, reclaim, prior_lo, vision_md)

    if save_docx:
        final_output = output_path or OUTPUT_DIR / f"{ticker}_analysis_{datetime.now().strftime('%Y%m%d')}.docx"
        stock_analyzer.generate_report(
            ticker, df, bull, bear, action, tag, reclaim, prior_lo, chart_path, vision_md, str(final_output)
        )
        output_path = final_output

    sections = build_render_sections(
        ticker, tag, structure, stage_info, vcp, kell, fib, vah, action, bull, bear, ticker_obj
    )

    return {
        "ticker": ticker,
        "tag": tag,
        "output": buffer.getvalue(),
        "docx_path": str(output_path) if output_path and Path(output_path).exists() else None,
        "last_price": f"${df.iloc[-1]['Close']:.2f}",
        "consult_time": datetime.now().strftime("%b %d, %Y %I:%M %p PT"),
        **sections,
    }


@app.get("/health")
def health():
    return jsonify({"ok": True})


@app.route("/", methods=["GET", "POST"])
def index():
    context = {
        "ticker": "",
        "output": "",
        "error": "",
        "download_url": "",
        "show_live": False,
        "chart_symbol": "",
        "consult_time": "",
        "insight_text": "",
        "stage_text": "",
        "volatility_text": "",
        "earnings_text": "",
        "action_text": "",
        "save_docx": False,
    }
    temp_chart_path = None

    if request.method == "POST":
        ticker = request.form.get("ticker", "").strip().upper()
        action = request.form.get("action", "consult")
        chart = request.files.get("chart")
        save_docx = action == "export"
        context["ticker"] = ticker
        context["save_docx"] = save_docx

        if not ticker:
            context["error"] = "Enter a ticker before consulting The Oracle."
            return render_template_string(INDEX_HTML, **context)

        if chart and chart.filename:
            suffix = Path(secure_filename(chart.filename)).suffix or ".png"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                chart.save(tmp)
                temp_chart_path = tmp.name

        try:
            result = run_analysis(ticker, "3y", temp_chart_path, save_docx)
            context.update(result)
            context["show_live"] = True
            context["chart_symbol"] = ticker
            if result["docx_path"]:
                context["download_url"] = f"/download/{Path(result['docx_path']).name}"
        except Exception as exc:
            context["error"] = f"{type(exc).__name__}: {exc}"
            context["output"] = context["error"]
        finally:
            if temp_chart_path and os.path.exists(temp_chart_path):
                os.unlink(temp_chart_path)

    return render_template_string(INDEX_HTML, **context)


@app.post("/api/analyze")
def analyze():
    ticker = request.form.get("ticker", request.args.get("ticker", "")).strip().upper()
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
        result = run_analysis(ticker, "3y", temp_chart_path, save_docx)
        payload = {
            "ok": True,
            "ticker": result["ticker"],
            "tag": result["tag"],
            "output": result["output"],
            "insight": result["insight_text"],
            "stage": result["stage_text"],
            "volatility": result["volatility_text"],
            "earnings": result["earnings_text"],
            "action": result["action_text"],
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
    return send_from_directory(str(OUTPUT_DIR), filename, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
