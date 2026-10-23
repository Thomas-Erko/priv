import re, shutil
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_from_directory

from report import build_report
import re
import pandas as pd
import dr8_report

ROOT = Path(__file__).parent
UPLOADS, OUTPUT = ROOT / "uploads", ROOT / "output"
UPLOADS.mkdir(exist_ok=True)
OUTPUT.mkdir(exist_ok=True)
ALLOWED = {".xlsx", ".xls", ".csv"}

app = Flask(__name__)


def safe(name: str) -> str:
    return re.sub(r"[^\w.\- ()一-鿿]", "_", Path(name).name)


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/upload")
def upload():
    saved = []
    for f in request.files.getlist("files"):
        name = safe(f.filename)
        if Path(name).suffix.lower() in ALLOWED and not name.startswith("~$"):
            f.save(UPLOADS / name)
            saved.append(name)
    return jsonify(saved=saved)


@app.get("/files")
def files():
    return jsonify(files=[dict(name=p.name, size=p.stat().st_size)
                          for p in sorted(UPLOADS.iterdir()) if p.suffix.lower() in ALLOWED])


@app.post("/clear")
def clear():
    shutil.rmtree(UPLOADS)
    UPLOADS.mkdir()
    return jsonify(ok=True)


@app.post("/generate")
def generate():
    body = request.get_json()
    results = []
    for name in body["files"]:
        p = UPLOADS / safe(name)
        try:
            if body.get("kind") == "dr8":
                out = OUTPUT / p.stem / f"{p.stem}_DR8_Data_Distribution.pptx"
                out.parent.mkdir(exist_ok=True)
                cfg = dr8_report.load_config(None)
                df = dr8_report.prep(dr8_report.load_frames([str(p)]), cfg)
                rx = re.compile(cfg["scope_process_regex"], re.I) if cfg.get("scope_process_regex") else None
                procs = [x for x in df["_process"].unique() if rx and rx.search(x)]
                if procs:
                    df = df[df["_process"].isin(procs)]
                if df.empty:
                    raise ValueError("No rows left after the final-test process filter")
                dr8_report.build_deck(df.reset_index(drop=True), cfg, str(out), str(ROOT / cfg["template"]))
                results.append(dict(file=name, url=f"/output/{p.stem}/{out.name}", path=str(out.relative_to(ROOT))))
                continue
            out = build_report(p, OUTPUT / p.stem, group=body.get("group") or None,
                               kind=body.get("kind", "box"), max_charts=int(body.get("max", 200)))
            results.append(dict(file=name, url=f"/output/{p.stem}/report.html", path=str(out.relative_to(ROOT))))
        except (Exception, SystemExit) as e:
            results.append(dict(file=name, error=str(e)))
    return jsonify(results=results)


@app.get("/output/<path:p>")
def output(p):
    return send_from_directory(OUTPUT, p)


if __name__ == "__main__":
    app.run(port=5055, debug=False)
