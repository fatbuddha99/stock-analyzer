from __future__ import annotations

import io
import os
import tempfile
from contextlib import redirect_stdout
from datetime import datetime
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
  <title>The Oracle</title>
  <style>
    :root {
      --bg: #030805;
      --panel: rgba(7, 18, 10, 0.86);
      --panel-2: rgba(10, 24, 14, 0.94);
      --line: rgba(115, 255, 143, 0.18);
      --text: rgba(236, 242, 236, 0.82);
      --muted: #7ec58b;
      --green: #7dff8b;
      --ink: #011006;
      --blue: #8df5ff;
      --red: #ff6f91;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      color: var(--text);
      background:
        radial-gradient(circle at top, rgba(61, 255, 136, 0.08), transparent 28%),
        radial-gradient(circle at 80% 10%, rgba(141, 245, 255, 0.05), transparent 18%),
        linear-gradient(180deg, #020503 0%, #050b06 42%, #020503 100%);
      font-family: "Courier New", Courier, monospace;
    }
    body::before {
      content: "";
      position: fixed;
      inset: 0;
      background:
        linear-gradient(rgba(125, 255, 139, 0.03) 1px, transparent 1px),
        linear-gradient(90deg, rgba(125, 255, 139, 0.02) 1px, transparent 1px);
      background-size: 100% 4px, 4px 100%;
      pointer-events: none;
      opacity: 0.3;
    }
    .shell {
      position: relative;
      max-width: 760px;
      margin: 0 auto;
      padding: 18px 14px 44px;
    }
    .topbar {
      display: grid;
      gap: 10px;
      margin-bottom: 14px;
    }
    .brand-row, .nav-row, .action-row, .card {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 18px;
      box-shadow:
        0 0 0 1px rgba(20, 70, 32, 0.18) inset,
        0 18px 55px rgba(0, 0, 0, 0.42);
      backdrop-filter: blur(8px);
    }
    .brand-row {
      padding: 16px 18px;
      background: linear-gradient(180deg, rgba(17, 36, 22, 0.95), rgba(6, 16, 9, 0.94));
    }
    .brand-row h1 {
      margin: 0;
      color: #020503;
      font-size: 1.7rem;
      font-weight: 700;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      font-family: Impact, Haettenschweiler, "Arial Narrow Bold", sans-serif;
      -webkit-text-stroke: 1.5px var(--green);
      text-shadow: 0 0 14px rgba(125, 255, 139, 0.18);
    }
    .brand-row p {
      margin: 8px 0 0;
      color: var(--muted);
      font-size: 0.8rem;
      line-height: 1.45;
    }
    .nav-row, .action-row {
      display: flex;
      gap: 8px;
      overflow-x: auto;
      flex-wrap: nowrap;
      padding: 8px;
      align-items: center;
    }
    .nav-row::-webkit-scrollbar, .action-row::-webkit-scrollbar { display: none; }
    .btn {
      flex: 0 0 auto;
      border: 1px solid rgba(240, 244, 240, 0.18);
      border-radius: 999px;
      background: rgba(232, 238, 232, 0.08);
      color: rgba(245, 248, 245, 0.8);
      padding: 7px 10px;
      font-size: 0.67rem;
      cursor: pointer;
      white-space: nowrap;
      text-transform: uppercase;
      letter-spacing: 0.09em;
      font-family: Verdana, Geneva, sans-serif;
      text-decoration: none;
    }
    .consult-btn {
      color: var(--green);
      border-color: rgba(125, 255, 139, 0.28);
      background: rgba(32, 75, 40, 0.22);
      font-family: "Trebuchet MS", Helvetica, sans-serif;
      font-weight: 700;
      letter-spacing: 0.12em;
    }
    .ticker-input {
      flex: 0 0 112px;
      min-width: 112px;
      border: 1px solid rgba(240, 244, 240, 0.18);
      border-radius: 999px;
      background: rgba(232, 238, 232, 0.05);
      color: rgba(245, 248, 245, 0.88);
      padding: 7px 12px;
      font: inherit;
      font-size: 0.72rem;
      font-family: Verdana, Geneva, sans-serif;
      text-transform: uppercase;
      outline: none;
    }
    .ticker-input::placeholder { color: rgba(245, 248, 245, 0.42); }
    .hidden-input { display: none; }
    .main { display: grid; gap: 12px; }
    .card { padding: 16px; }
    .hero {
      display: grid;
      gap: 12px;
      background: linear-gradient(180deg, rgba(8, 20, 11, 0.95), rgba(5, 12, 7, 0.95));
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
      color: #ffb347;
    }
    .dot {
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: #ffb347;
      box-shadow: 0 0 10px rgba(125, 255, 139, 0.55);
    }
    .hero p, .card p {
      margin: 0;
      line-height: 1.6;
      font-size: 0.91rem;
    }
    .chart-card {
      padding: 0;
      overflow: hidden;
      background: var(--panel-2);
    }
    .chart-wrap { height: 420px; background: #050b06; }
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 10px;
    }
    .kpi {
      padding: 12px;
      border: 1px solid var(--line);
      border-radius: 14px;
      background: rgba(8, 18, 10, 0.86);
    }
    .kpi .label {
      color: var(--muted);
      font-size: 0.7rem;
      margin-bottom: 8px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      font-family: Arial, Helvetica, sans-serif;
    }
    .kpi .value {
      font-size: 1.08rem;
      font-weight: 700;
      line-height: 1.2;
      color: #f4fff5;
    }
    .section-title h2 {
      margin: 0 0 10px;
      font-size: 0.98rem;
      color: var(--green);
      font-family: "Trebuchet MS", Helvetica, sans-serif;
      letter-spacing: 0.03em;
    }
    .error { color: var(--red); }
    pre {
      margin: 0;
      white-space: pre-wrap;
      overflow-x: auto;
      color: rgba(236, 242, 236, 0.82);
      font: inherit;
      line-height: 1.55;
    }
    .download-link { color: var(--blue); }
    @media (max-width: 640px) {
      .shell { padding: 12px 10px 32px; }
      .ticker { font-size: 1.6rem; }
      .ticker-row { flex-direction: column; align-items: flex-start; }
      .meta { text-align: left; }
      .chart-wrap { height: 360px; }
      .kpi-grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <div class="shell">
    <form method="post" enctype="multipart/form-data" id="oracle-form">
      <input class="hidden-input" id="chart-input" type="file" name="chart" accept=".png,.jpg,.jpeg,.webp">
      <input type="hidden" name="docx" value="{{ '1' if save_docx else '0' }}">
      <input type="hidden" name="period" value="3y">
      <div class="topbar">
        <div class="brand-row">
          <h1>The Oracle</h1>
          <p>A Deep Dive For Context And Nuances.</p>
        </div>
        <div class="nav-row">
          <a class="btn" href="#insight">Insight</a>
          <a class="btn" href="#stage">Stage</a>
          <a class="btn" href="#volatility">Volatility</a>
          <a class="btn" href="#earnings">Earnings</a>
          <a class="btn" href="#tailwind">Tailwind</a>
          <a class="btn" href="#action">Action</a>
          <button class="btn" type="button" id="upload-trigger">Upload</button>
          <button class="btn" type="submit" name="action" value="export">Export</button>
        </div>
        <div class="action-row">
          <input class="ticker-input" name="ticker" value="{{ ticker }}" placeholder="Ticker" aria-label="Ticker" required>
          <button class="btn consult-btn" type="submit" name="action" value="consult">Consult</button>
        </div>
      </div>
    </form>

    <div class="main">
      <div class="card hero">
        <div class="verdict"><span class="dot"></span> {% if show_live %}Oracle Consult{% else %}Stand By{% endif %}</div>
        <div class="ticker-row">
          <div class="ticker">{% if show_live %}{{ ticker }}{% else %}The Oracle Is Self Fufilling{% endif %}</div>
          <div class="meta">{% if show_live %}Apr 28, 2026 1:15 PM PT{% else %}Waiting for a ticker consult{% endif %}</div>
        </div>
        {% if show_live %}
        <p>The Oracle has generated a live consult for {{ ticker }}. Use the chart, KPI row, and Oracle Output below as the active read for this ticker instead of the default sample narrative.</p>
        {% else %}
        <p>Type a ticker, then press Consult. The landing page stays intentionally blank and lightweight until you ask for a read, so the chart and analysis only load when they are actually needed.</p>
        {% endif %}
      </div>

      {% if show_live %}
      <div class="card chart-card">
        <div class="chart-wrap">
          <iframe
            title="TradingView Chart"
            src="https://s.tradingview.com/widgetembed/?symbol=NASDAQ%3A{{ chart_symbol }}&interval=D&hidesidetoolbar=1&symboledit=1&saveimage=0&toolbarbg=050b06&studies=[]&theme=dark&style=1&timezone=Etc%2FUTC&withdateranges=1&hideideas=1&hidevolume=0&calendar=0&details=0&hotlist=0&news=0&watchlist=0&locale=en"
            style="width: 100%; height: 100%; border: 0;"
            allowtransparency="true"
            scrolling="no"></iframe>
        </div>
      </div>

      <div class="card">
        <div class="kpi-grid">
          <div class="kpi">
            <div class="label">Current Price</div>
            <div class="value">{{ last_price or '$38.42' }}</div>
          </div>
          <div class="kpi">
            <div class="label">Last Earnings</div>
            <div class="value">{{ earnings_blurb or 'EPS Beat · Feb 10' }}</div>
          </div>
          <div class="kpi">
            <div class="label">Reaction</div>
            <div class="value">{{ reaction_blurb or 'Revenue Miss Fade' }}</div>
          </div>
        </div>
      </div>
      {% endif %}

      {% if error %}
      <div class="card">
        <div class="section-title"><h2>Error</h2></div>
        <p class="error">{{ error }}</p>
      </div>
      {% endif %}

      {% if output %}
      <div class="card" id="consult-output">
        <div class="section-title"><h2>Oracle Output</h2></div>
        <pre>{{ output }}</pre>
        {% if download_url %}
        <p style="margin-top: 12px;">DOCX created: <a class="download-link" href="{{ download_url }}">{{ download_url }}</a></p>
        {% endif %}
      </div>
      {% endif %}

      {% if not show_live %}
      <div class="card" id="insight">
        <div class="section-title"><h2>Insight</h2></div>
        <p>The Oracle does not know everything. It waits for a ticker, then builds a live consult from the chart, recent structure, earnings context, and the script's own analysis before it says anything specific.</p>
      </div>
      {% endif %}
    </div>
  </div>
  <script>
    const uploadTrigger = document.getElementById('upload-trigger');
    const chartInput = document.getElementById('chart-input');
    uploadTrigger?.addEventListener('click', () => chartInput?.click());
  </script>
</body>
</html>
"""


def run_analysis(ticker: str, period: str, chart_path: str | None, save_docx: bool):
    ticker = ticker.upper().strip()
    output_dir = Path("outputs")
    output_dir.mkdir(exist_ok=True)
    output_path = output_dir / f"{ticker}_analysis.docx" if save_docx else None

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
        final_output = output_path or output_dir / f"{ticker}_analysis_{datetime.now().strftime('%Y%m%d')}.docx"
        stock_analyzer.generate_report(
            ticker,
            df,
            bull,
            bear,
            action,
            tag,
            reclaim,
            prior_lo,
            chart_path,
            vision_md,
            str(final_output),
        )
        output_path = final_output

    return {
        "ticker": ticker,
        "tag": tag,
        "output": buffer.getvalue(),
        "docx_path": str(output_path) if output_path and Path(output_path).exists() else None,
        "last_price": f"${df.iloc[-1]['Close']:.2f}",
        "action_plan": action,
        "reclaim_tier1": reclaim["tier1"]["price"],
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
    action = "consult"
    last_price = "$38.42"
    earnings_blurb = "EPS Beat · Feb 10"
    reaction_blurb = "Revenue Miss Fade"
    action_plan = ("Watch for a tighter post-earnings base and a move that actually sticks above "
                   "the next reclaim zone above $40.20.")

    if request.method == "POST":
        ticker = request.form.get("ticker", "").strip().upper()
        period = request.form.get("period", "3y")
        action = request.form.get("action", "consult")
        save_docx = action == "export" or request.form.get("docx", "0") == "1"
        chart = request.files.get("chart")

        if not ticker:
            error = "Enter a ticker before consulting The Oracle."
            return render_template_string(
                INDEX_HTML,
                output=output,
                error=error,
                ticker=ticker,
                chart_symbol="",
                period=period,
                save_docx=save_docx,
                download_url=download_url,
                last_price=last_price,
                earnings_blurb=earnings_blurb,
                reaction_blurb=reaction_blurb,
                action_plan=action_plan,
                periods=["1y", "2y", "3y", "5y", "10y", "max"],
            )

        if chart and chart.filename:
            suffix = Path(secure_filename(chart.filename)).suffix or ".png"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                chart.save(tmp)
                temp_chart_path = tmp.name

        try:
            result = run_analysis(ticker, period, temp_chart_path, save_docx)
            output = result["output"]
            last_price = result["last_price"]
            action_plan = result["action_plan"]
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
        chart_symbol=ticker,
        show_live=bool(ticker and output),
        period=period,
        save_docx=save_docx,
        download_url=download_url,
        last_price=last_price,
        earnings_blurb=earnings_blurb,
        reaction_blurb=reaction_blurb,
        action_plan=action_plan,
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
