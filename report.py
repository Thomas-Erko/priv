import base64, html, io, re
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams["font.sans-serif"] = ["PingFang SC", "Heiti SC", "Arial Unicode MS", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False
import pandas as pd

LANE_RE = re.compile(r"^(.*?)[_ ]?(?:Lane|CH|Ch|ch)(\d+)$")
GROUP_CANDIDATES = ["PROCESS_NAME", "STEP_ID", "STATION_NUMBER", "RESULT"]


def load_sheets(path: Path) -> dict[str, pd.DataFrame]:
    if path.suffix.lower() == ".csv":
        for enc in ("utf-8-sig", "gbk", "latin1"):
            try:
                return {path.stem: pd.read_csv(path, encoding=enc, low_memory=False)}
            except UnicodeDecodeError:
                continue
    return pd.read_excel(path, sheet_name=None)


def numeric_columns(df: pd.DataFrame) -> dict[str, pd.Series]:
    out = {}
    for c in df.columns:
        s = pd.to_numeric(df[c], errors="coerce")
        nn = df[c].notna().sum()
        if nn >= 2 and s.notna().sum() >= 0.5 * nn and s.nunique() > 1:
            out[c] = s
    return out


def fig_to_img(fig) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=80, bbox_inches="tight")
    plt.close(fig)
    return '<img src="data:image/png;base64,%s">' % base64.b64encode(buf.getvalue()).decode()


def plot_metric(name, cols, df, num, group, kind):
    """cols: {label: column}. One box (or hist) per label, one subplot per group value."""
    groups = [None]
    if group:
        counts = {g: sum(num[c][df[group] == g].notna().sum() for c in cols.values()) for g in df[group].dropna().unique()}
        groups = [g for g, n in sorted(counts.items(), key=lambda kv: -kv[1]) if n >= 2][:6] or [None]
    fig, axes = plt.subplots(len(groups), 1, figsize=(8, 2.6 * len(groups) + 0.6), squeeze=False)
    drew = False
    for ax, g in zip(axes[:, 0], groups):
        mask = (df[group] == g) if g is not None else slice(None)
        data = [(lbl, num[c][mask].dropna()) for lbl, c in cols.items()]
        data = [(l, d) for l, d in data if len(d)]
        if not data:
            ax.set_visible(False)
            continue
        drew = True
        if kind == "hist":
            for l, d in data:
                ax.hist(d, bins=20, alpha=0.5, label=l)
            if len(data) > 1:
                ax.legend(fontsize=7, ncol=4)
        else:
            ax.boxplot([d for _, d in data], tick_labels=[l for l, _ in data])
        if g is not None:
            ax.set_title(str(g), fontsize=9, loc="right")
        ax.grid(alpha=0.3)
    if not drew:
        plt.close(fig)
        return None
    fig.suptitle(name, fontweight="bold")
    fig.tight_layout()
    return fig_to_img(fig)


def build_report(path: Path, out_dir: Path, group=None, kind="box", max_charts=200) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    parts = [f"<h1>{html.escape(path.name)}</h1>"]
    for sheet, df in load_sheets(path).items():
        df = df.dropna(how="all")
        num = numeric_columns(df)
        g = group if group in df.columns else next((c for c in GROUP_CANDIDATES if c in df.columns and c != "RESULT"), None)
        parts.append(f"<h2>Sheet: {html.escape(str(sheet))}</h2><p>{len(df)} rows, {len(df.columns)} columns, {len(num)} numeric. Grouped by: <b>{html.escape(str(g))}</b></p>")
        if "RESULT" in df.columns:
            parts.append("<h3>Result counts</h3>" + df["RESULT"].value_counts().to_frame().to_html())
        # collapse *_LaneN columns into one multi-box chart
        metrics: dict[str, dict[str, str]] = {}
        for c in num:
            m = LANE_RE.match(str(c))
            key, label = (m.group(1) or str(c), "L" + m.group(2)) if m else (str(c), str(c))
            metrics.setdefault(key, {})[label] = c
        parts.append("<h3>Charts</h3>")
        n = 0
        for key, cols in metrics.items():
            if n >= max_charts:
                parts.append(f"<p>Stopped at {max_charts} charts.</p>")
                break
            img = plot_metric(key, cols, df, num, g, kind)
            if img:
                parts.append(f'<div class="c">{img}</div>')
                n += 1
        if num:
            summ = pd.DataFrame(num).describe().T[["count", "mean", "std", "min", "max"]].round(4)
            parts.append("<h3>Summary stats</h3>" + summ.to_html(classes="t"))
    page = ("<!doctype html><meta charset=utf-8><title>%s</title><style>body{font:14px sans-serif;margin:24px;max-width:1100px}"
            ".c{display:inline-block;margin:6px;vertical-align:top}table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:2px 6px}</style>%s"
            % (html.escape(path.name), "".join(parts)))
    out = out_dir / "report.html"
    out.write_text(page, encoding="utf-8")
    return out
