#!/usr/bin/env python3
"""
make_dr8_report.py
Builds the "DR8 Data Distribution" deck (JMP-style variability charts + capability
pages) straight from station test-data exports.

Quick start
  python make_dr8_report.py                          # asks for everything
  python make_dr8_report.py data.xlsx -o DR8_report.pptx
  python make_dr8_report.py a.xlsx b.xlsx --product DR8 --process FinalTest
  python make_dr8_report.py --dump-config dr8_config.json    # edit limits/charts
  python make_dr8_report.py data.xlsx --config dr8_config.json

Needs: pandas, openpyxl, python-pptx, matplotlib, scipy
Keep dr8_template.pptx next to this script (or pass --template).

Slide recipe (all driven by the config):
  1   title slide          "<product> Data Distribution"
  2-3 Performance Distribution overview     variability charts, one panel per test temperature
  4+  capability pages     normal fit at mean±3σ and mean±2σ, like JMP's
                           Distribution > Capability Analysis output
"""

import argparse
import copy
import json
import math
import os
import re
import sys
import tempfile

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.patches import Polygon, Rectangle
from matplotlib.ticker import FuncFormatter, LogLocator, MaxNLocator, NullFormatter
from pptx import Presentation
from pptx.util import Inches
from scipy import stats

HERE = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------------------
# Config. --dump-config writes this out so limits/columns can be edited.
# Limits mirror the DR8 reference deck; confirm against the released spec.
# ---------------------------------------------------------------------------
DEFAULT_CONFIG = {
    "product": "DR8",
    "template": "dr8_template.pptx",
    "columns": {
        "serial": "SERIAL_NUMBER",
        "process": "PROCESS_NAME",
        "start": "START_DATE_TIME",
        "result": "RESULT",
        "temperature": ["TemperatureSetting", "CaseTemperature"],
    },
    # default scope when no --process is given: processes whose name matches
    "scope_process_regex": "final",
    "temperature_round_to": 5,
    "temp_labels": {"low": "Low Temp (LT Testing)", "room": "Room Temp (RT Testing)",
                    "high": "High Temp (HT Testing)"},
    # capability fits use passing runs only; failed runs (aborted / unpowered
    # measurements) would otherwise blow up sigma. Overview charts keep them.
    "capability_exclude_failed": True,
    "pass_values": ["PASS", "OK", "P", "PASSED", "GOOD"],
    "invalid_abs_value": 1e9,
    "lane_regex": r"^(?P<base>.+?)_Lane(?P<lane>\d+)$",

    # kind: lanes      -> CH1..CHn boxes, one stacked panel per temperature
    #       category   -> one box per temperature (scalar column)
    #       power_drop -> per unit & lane, max-min of `col` across temperatures
    # col may be a list: the first column with non-constant data wins.
    "charts": {
        "TDSP": {"title": "TDSP", "kind": "category", "col": ["DspTemperature", "LutTemperature"],
                 "ylabel": "TDSP", "tick_fmt": "%d"},
        "PowerConsumption": {"title": "Power Consumption", "kind": "category", "col": ["PowerConsumption"],
                             "ylabel": "Power Consumption(W)", "ucl": 16},
        "TxPower": {"title": "Tx Power", "kind": "lanes", "col": ["TxPower"], "ylabel": "Pave(dBm)", "lcl": -2.4},
        "OMA": {"title": "OMA", "kind": "lanes", "col": ["TxOMA"], "ylabel": "OMA(dBm)", "ucl": 4, "lcl": -0.5},
        "OMA_TDECQ": {"title": "OMA-TDECQ", "kind": "lanes", "col": ["TxOMA_Minus_TDECQ"],
                      "ylabel": "OMA-TDECQ(dB)", "lcl": -1.8},
        "TDECQ": {"title": "TDECQ", "kind": "lanes", "col": ["TDECQ"], "ylabel": "TDECQ(dB)", "ucl": 3.4,
                  "tick_fmt": "%d"},
        "ER": {"title": "ER", "kind": "lanes", "col": ["OutER", "ER"], "ylabel": "ER(dB)", "ucl": 8, "lcl": 4},
        "Sensitivity": {"title": "Sensitivity", "kind": "lanes", "col": ["Sensitivity"],
                        "ylabel": "Sensitivity(dBm)", "ucl": -6},
        "Overload": {"title": "Overload", "kind": "lanes", "col": ["BER_Overload"], "ylabel": "BER",
                     "ucl": 1e-8, "log": True},
        "Wavelength": {"title": "Wavelength", "kind": "lanes", "col": ["Wavelength"], "ylabel": "Wavelength(nm)",
                       "ucl": 1317, "lcl": 1305, "tick_fmt": "%d"},
        "BiasCurrent": {"title": "Bias Current", "kind": "lanes", "col": ["BiasCurrent"],
                        "ylabel": "Bias Current(mA)"},
        "RxPowerDelta": {"title": "RxPower Delta@Input power={setting}dB", "kind": "lanes",
                         "col_by_setting": {"delta": "DDMI_RxPower{n}_Delta", "setting": "DDMI_RxPower{n}_Setting",
                                            "target": -13},
                         "ylabel": "", "ucl": 1.5, "lcl": -1.5},
        "TxPowerDelta": {"title": "TxPower Delta", "kind": "lanes", "col": ["DDMI_TxPower_Delta"],
                         "ylabel": "", "ucl": 1.5, "lcl": -1.5},
        "PowerDrop": {"title": "Tri-Temp Power Drop", "kind": "power_drop", "col": ["TxPower"], "ylabel": "delta"},
    },
    # overview slides: chart key + placement in inches on the 13.33 x 7.5 slide
    "overview_slides": [
        {"title": "Performance Distribution", "w": 3.54, "h": 2.76, "charts": [
            ["TDSP", 0.04, 1.12], ["PowerConsumption", 3.27, 1.12], ["TxPower", 6.49, 1.12], ["OMA", 9.71, 1.12],
            ["OMA_TDECQ", 0.04, 4.15], ["TDECQ", 3.30, 4.15], ["ER", 6.56, 4.15], ["Sensitivity", 9.71, 4.15]]},
        {"title": "Performance Distribution", "w": 3.94, "h": 3.15, "charts": [
            ["Overload", 0.25, 1.01], ["Wavelength", 4.51, 1.10], ["BiasCurrent", 8.70, 1.10],
            ["RxPowerDelta", 0.21, 4.25], ["TxPowerDelta", 4.44, 4.20], ["PowerDrop", 8.65, 4.20]]},
    ],
    # capability pages: side = upper | lower | both, temp = high to use HT only
    "capability_slides": [
        {"chart": "TDSP", "slide_title": "HT TDSP", "name": "TDSP", "temp": "high", "side": "upper"},
        {"chart": "PowerConsumption", "slide_title": "Power Consumption", "name": "Power Consumption",
         "temp": "high", "side": "upper"},
        {"chart": "TxPower", "slide_title": "Tx power", "name": "Tx power", "side": "both"},
        {"chart": "OMA", "slide_title": "OMA", "name": "OMA", "side": "both"},
        {"chart": "OMA_TDECQ", "slide_title": "OMA-TDECQ", "name": "OMA-TDECQ", "side": "lower"},
        {"chart": "TDECQ", "slide_title": "TDECQ", "name": "TDECQ", "side": "upper"},
        {"chart": "Sensitivity", "slide_title": "Sensitivity", "name": "Sensitivity", "side": "upper"},
        {"chart": "PowerDrop", "slide_title": "Tri-Temp Power Drop", "name": "Tri-Temp Power Drop", "side": "upper"},
    ],
    "sigma_levels": [3, 2],
    "capability_placement": {"x": 0.22, "w": 11.03, "h": 3.02, "y_top": 1.15, "y_bottom": 4.23},
}

# JMP-ish look
RED, LIMIT_RED, BLUE = "#E0465A", "#E8384F", "#3A5FCD"
FRAME, STRIP, HDR_BG, HDR_EC, LABEL_BG = "#A0A0A0", "#E4E4E4", "#EFEFEF", "#CFCFCF", "#E9E9E9"
BAR, BAR_EDGE = "#B3C2AE", "#55634F"
plt.rcParams.update({
    "font.family": "sans-serif",
    "font.sans-serif": ["Microsoft YaHei", "Segoe UI", "Noto Sans CJK SC", "SimHei", "PingFang SC",
                        "Arial Unicode MS", "DejaVu Sans"],
    "axes.unicode_minus": False,
    "font.size": 8,
})


# ============================== data =======================================

def load_frames(paths):
    frames = []
    for p in paths:
        if os.path.splitext(p)[1].lower() in (".csv", ".txt"):
            frames.append(pd.read_csv(p, low_memory=False))
        else:
            for _, sh in pd.read_excel(p, sheet_name=None).items():
                if len(sh.columns) > 5 and len(sh):
                    frames.append(sh)
    if not frames:
        sys.exit("No data rows found in the input file(s).")
    return pd.concat(frames, ignore_index=True, sort=False)


CJK = re.compile(r"[\u3000-\u303f\u3040-\u30ff\u3400-\u9fff\uff00-\uffef]+")
ZH_EN = {"老化前测试": "Before Aging Test", "老化后测试": "After Aging Test", "高温测试": "HT Testing",
         "常温测试": "RT Testing", "低温测试": "LT Testing", "标定": "Calibration", "温循传纤测试": "Temp Cycle",
         "三温": "Tri-Temp", "数据": "Data"}


def to_english(text):
    """Keep output English: use the English part of 'CN (EN)' labels, translate known
    terms, and drop any other CJK characters."""
    s = str(text)
    if not CJK.search(s):
        return s
    m = re.search(r"\(([^()]*[A-Za-z][^()]*)\)", s)
    if m and not CJK.search(m.group(1)):
        return m.group(1).strip()
    for zh, en in sorted(ZH_EN.items(), key=lambda kv: -len(kv[0])):
        s = s.replace(zh, f" {en} ")
    s = re.sub(r"\s+", " ", CJK.sub(" ", s)).strip(" -_,;:")
    return s or "Unlabeled"


def prep(df, cfg):
    c = cfg["columns"]
    if c["serial"] not in df.columns:
        sys.exit(f"Required column '{c['serial']}' not found. Edit the config 'columns' block.")
    out = df.copy()
    out["_serial"] = out[c["serial"]].astype(str).str.strip()
    out["_process"] = (out[c["process"]].fillna("Unknown").astype(str).map(to_english)
                       if c["process"] in out else "Unknown")
    out["_start"] = pd.to_datetime(out[c["start"]], errors="coerce") if c["start"] in out else pd.NaT
    passes = {v.upper() for v in cfg.get("pass_values", ["PASS"])}
    out["_pass"] = (out[c["result"]].astype(str).str.strip().str.upper().isin(passes)
                    if c["result"] in out else True)
    tcol = next((n for n in c["temperature"] if n in out.columns), None)
    if tcol is None:
        sys.exit("No temperature column found (TemperatureSetting / CaseTemperature).")
    step = cfg["temperature_round_to"] or 1
    out["_temp"] = (pd.to_numeric(out[tcol], errors="coerce") / step).round() * step
    return out.sort_values("_start", kind="stable").reset_index(drop=True)


def temp_label(t, cfg):
    lab = cfg["temp_labels"]
    if t < 15:
        return lab["low"]
    if t > 50:
        return lab["high"]
    return lab["room"]


def temp_groups(df):
    """Sorted temperatures present (LT, RT, HT order)."""
    return sorted(df["_temp"].dropna().unique())


def numeric(s, cfg, log=False):
    v = pd.to_numeric(s, errors="coerce")
    v = v[v.abs() < cfg["invalid_abs_value"]].dropna()
    return v[v > 0] if log else v


def lane_cols(df, base, cfg):
    rx = re.compile(cfg["lane_regex"])
    found = []
    for col in df.columns:
        m = rx.match(str(col))
        if m and m.group("base") == base:
            found.append((int(m.group("lane")), col))
    found.sort()
    start = 1 if found and found[0][0] == 0 else 0
    return [(f"CH{n + start}", c) for n, c in found]


def resolve_columns(df, spec, cfg):
    """Pick the source column(s) for a chart. Returns (cols, info) where cols is
    [(lane_label or None, column)] and info carries notes like the setting used."""
    info = {}
    if "col_by_setting" in spec:
        cb = spec["col_by_setting"]
        best = None
        for n in range(1, 17):
            sc = lane_cols(df, cb["setting"].format(n=n), cfg)
            dc = lane_cols(df, cb["delta"].format(n=n), cfg)
            if not sc or not dc:
                continue
            setting = numeric(df[sc[0][1]], cfg).median()
            vals = pd.concat([numeric(df[c], cfg) for _, c in dc])
            if pd.isna(setting) or vals.nunique() <= 1:
                continue
            if best is None or abs(setting - cb["target"]) < abs(best[0] - cb["target"]):
                best = (setting, dc)
        if best is None:
            return [], info
        info["setting"] = f"{best[0]:g}"
        return best[1], info

    first = None
    for base in spec["col"]:
        if spec["kind"] in ("lanes", "power_drop"):
            cols = lane_cols(df, base, cfg)
        else:
            cols = [(None, base)] if base in df.columns else []
        if not cols:
            continue
        vals = pd.concat([numeric(df[c], cfg, spec.get("log")) for _, c in cols])
        first = first or (cols, base)
        if vals.nunique() > 1:
            if base != spec["col"][0]:
                info["note"] = f"{spec['col'][0]} is empty/constant, using {base}"
            info["base"] = base
            return cols, info
    if first:
        info["note"] = f"{first[1]} has no usable spread"
        info["base"] = first[1]
        return first[0], info
    return [], info


def chart_data(df, spec, cfg):
    """Returns dict with panels [(temp, label, [values per lane])], lane labels."""
    cols, info = resolve_columns(df, spec, cfg)
    title = spec["title"].format(**{"setting": info.get("setting", "?")})
    res = {"title": title, "info": info, "panels": [], "lanes": [lab for lab, _ in cols]}
    if not cols:
        return res
    temps = temp_groups(df)
    log = spec.get("log", False)

    if spec["kind"] == "power_drop":
        # latest run per unit & temperature, then max-min across temperatures per lane
        sub = df[df["_temp"].notna()]
        latest = sub.groupby(["_serial", "_temp"], sort=False).tail(1)
        per_lane = []
        for _, c in cols:
            v = pd.to_numeric(latest[c], errors="coerce")
            v = v.where(v.abs() < cfg["invalid_abs_value"])
            piv = latest.assign(_v=v).pivot_table(index="_serial", columns="_temp", values="_v", aggfunc="last")
            piv = piv.dropna(thresh=2)
            per_lane.append((piv.max(axis=1) - piv.min(axis=1)).dropna())
        res["panels"] = [(None, None, per_lane)]
        return res

    for t in temps:
        sub = df[df["_temp"] == t]
        if spec["kind"] == "category":
            v = numeric(sub[cols[0][1]], cfg, log)
            if len(v):
                res["panels"].append((t, temp_label(t, cfg), [v]))
        else:
            lanes = [numeric(sub[c], cfg, log) for _, c in cols]
            if any(len(v) for v in lanes):
                res["panels"].append((t, temp_label(t, cfg), lanes))
    return res


# ============================== JMP-style charts ===========================

def lim_text(tag, v):
    if v is not None and 0 < abs(v) < 1e-3:
        return f"{tag}={v:.0E}".replace("E-0", "E-").replace("E+0", "E+")
    return f"{tag}={v:g}"


def draw_boxes(ax, groups, positions, width=0.5):
    data = [np.asarray(v, dtype=float) for v in groups]
    keep = [i for i, v in enumerate(data) if len(v)]
    if not keep:
        return
    ax.boxplot([data[i] for i in keep], positions=[positions[i] for i in keep], widths=width,
               whis=1.5, patch_artist=True, manage_ticks=False, capwidths=width * 0.9,
               boxprops={"facecolor": "white", "edgecolor": "black", "linewidth": 1.0},
               medianprops={"color": "black", "linewidth": 1.0},
               whiskerprops={"color": "black", "linewidth": 1.0},
               capprops={"color": "black", "linewidth": 1.0},
               flierprops={"marker": "o", "markerfacecolor": "black", "markeredgecolor": "black",
                           "markersize": 3.4, "linestyle": "none"})


def y_range(values, spec, log=False):
    allv = pd.concat([pd.Series(v, dtype=float) for v in values if len(v)]) if values else pd.Series(dtype=float)
    lims = [spec.get(k) for k in ("ucl", "lcl") if spec.get(k) is not None]
    pts = list(allv) + lims
    if not pts:
        return (0, 1)
    lo, hi = min(pts), max(pts)
    if log:
        return (10 ** (math.floor(math.log10(lo)) - 0.15), 10 ** (math.log10(hi) + 0.35))
    if hi == lo:
        hi, lo = hi + 1, lo - 1
    pad = (hi - lo) * 0.09
    return (lo - pad, hi + pad)


def style_frame(ax):
    for sp in ax.spines.values():
        sp.set_color(FRAME)
        sp.set_linewidth(0.8)
    ax.tick_params(axis="y", length=3.5, width=0.7, color="#606060", labelsize=9.4, pad=2)
    ax.tick_params(axis="x", length=0, labelsize=9.4, pad=3)


def set_y_ticks(ax, spec, ylim, nbins, log=False):
    ax.set_ylim(*ylim)
    if log:
        ax.set_yscale("log")
        ax.yaxis.set_major_locator(LogLocator(base=10, numticks=7))
        ax.yaxis.set_minor_locator(LogLocator(base=10, subs=[], numticks=20))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"1e{int(round(math.log10(v)))}"))
        ax.yaxis.set_minor_formatter(NullFormatter())
    else:
        fmt = spec.get("tick_fmt", "%.1f")
        whole = fmt == "%d"
        ax.yaxis.set_major_locator(MaxNLocator(nbins=nbins, steps=[1, 2, 5, 10] if whole else [1, 2, 2.5, 5, 10],
                                               integer=whole))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: fmt % v))
    # keep tick labels off the panel edges so stacked panels don't collide
    span = (math.log10(ylim[1]) - math.log10(ylim[0])) if log else (ylim[1] - ylim[0])
    pos = lambda t: math.log10(t) if log else t
    lo_e, hi_e = pos(ylim[0]), pos(ylim[1])
    ax.set_yticks([t for t in ax.get_yticks() if ylim[0] <= t <= ylim[1]
                   and pos(t) - lo_e > span * 0.035 and hi_e - pos(t) > span * 0.035])
    # hide tick labels that would sit under a limit label, like JMP does
    lims = [spec.get(k) for k in ("ucl", "lcl") if spec.get(k) is not None]
    if lims:
        span = (math.log10(ylim[1]) - math.log10(ylim[0])) if log else (ylim[1] - ylim[0])
        ticks = [t for t in ax.get_yticks() if ylim[0] <= t <= ylim[1]]
        keep = []
        for t in ticks:
            d = min(abs((math.log10(t) - math.log10(l)) if log else (t - l)) for l in lims)
            if d > span * 0.08:
                keep.append(t)
        ax.set_yticks(keep)


def add_limits(fig, ax, spec, W):
    for key, tag in (("ucl", "UCL"), ("lcl", "LCL")):
        v = spec.get(key)
        if v is None:
            continue
        lo, hi = ax.get_ylim()
        if not (min(lo, hi) <= v <= max(lo, hi)):
            continue
        ax.axhline(v, color=LIMIT_RED, lw=1.0, ls=(0, (8, 2.5, 1.5, 2.5)), zorder=1)
        ax.annotate(lim_text(tag, v), xy=(0, v), xycoords=("axes fraction", "data"),
                    xytext=(-2, 0), textcoords="offset points", ha="right", va="center",
                    color=LIMIT_RED, fontsize=8.8, annotation_clip=False,
                    bbox={"fc": "white", "ec": "none", "pad": 0.3})


def variability_png(cd, spec, w_in, h_in, path):
    """Stacked-panel variability chart (lanes on x, one panel per temperature)."""
    W = 673
    H = round(W * h_in / w_in)
    fig = plt.figure(figsize=(W / 100, H / 100), dpi=100)
    left, right_strip, top, bottom = 66, 22, 32, 26
    x0, x1 = left, W - 103 - right_strip
    log = spec.get("log", False)
    panels = cd["panels"]
    fig.text((x0 + x1) / 2 / W, 1 - 15 / H, cd["title"], ha="center", va="center", fontsize=15.5, weight="bold")

    if not panels:
        ax = fig.add_axes([x0 / W, bottom / H, (x1 - x0) / W, 1 - (top + bottom) / H])
        style_frame(ax)
        ax.set_xticks([]), ax.set_yticks([])
        ax.text(0.5, 0.5, "No data in the selected scope", ha="center", va="center", color="#808080",
                transform=ax.transAxes)
        fig.savefig(path, dpi=200)
        plt.close(fig)
        return

    n = len(panels)
    ph = (H - top - bottom) / n
    lanes = cd["lanes"]
    ylim = y_range([v for _, _, vals in panels for v in vals], spec, log)
    nb = 7 if n > 1 else 9
    for i, (t, lab, vals) in enumerate(panels):
        yb = bottom + (n - 1 - i) * ph
        ax = fig.add_axes([x0 / W, yb / H, (x1 - x0) / W, ph / H])
        style_frame(ax)
        draw_boxes(ax, vals, list(range(1, len(vals) + 1)), 0.5)
        ax.set_xlim(0.4, len(lanes) + 0.6)
        set_y_ticks(ax, spec, ylim, nb, log)
        add_limits(fig, ax, spec, W)
        if i == n - 1:
            ax.set_xticks(range(1, len(lanes) + 1), lanes)
        else:
            ax.set_xticks([])
        # right-hand strip with the temperature name, rotated like JMP
        sx = fig.add_axes([x1 / W, yb / H, right_strip / W, ph / H])
        sx.set_facecolor(STRIP)
        sx.set_xticks([]), sx.set_yticks([])
        for sp in sx.spines.values():
            sp.set_color("#C8C8C8")
            sp.set_linewidth(0.6)
        sx.text(0.5, 0.5, lab, rotation=270, ha="center", va="center", fontsize=8.6, transform=sx.transAxes)
    if spec.get("ylabel"):
        fig.text(8 / W, (bottom + (H - top - bottom) / 2) / H, spec["ylabel"], rotation=90,
                 ha="center", va="center", fontsize=9.8)
    fig.savefig(path, dpi=200)
    plt.close(fig)


def single_png(cd, spec, w_in, h_in, path):
    """One-panel chart: boxes per temperature (category) or per lane (power drop)."""
    W = 673
    H = round(W * h_in / w_in)
    fig = plt.figure(figsize=(W / 100, H / 100), dpi=100)
    x0, x1, top, bottom = 55, W - 113, 32, 24
    ax = fig.add_axes([x0 / W, bottom / H, (x1 - x0) / W, (H - top - bottom) / H])
    style_frame(ax)
    fig.text((x0 + x1) / 2 / W, 1 - 15 / H, cd["title"], ha="center", va="center", fontsize=15.5, weight="bold")
    if spec["kind"] == "category":
        vals = [v for _, _, (v,) in cd["panels"]]
        labels = [lab for _, lab, _ in cd["panels"]]
    else:
        vals = cd["panels"][0][2] if cd["panels"] else []
        labels = cd["lanes"]
    if not vals or not any(len(v) for v in vals):
        ax.set_xticks([]), ax.set_yticks([])
        ax.text(0.5, 0.5, "No data in the selected scope", ha="center", va="center", color="#808080",
                transform=ax.transAxes)
    else:
        draw_boxes(ax, vals, list(range(1, len(vals) + 1)), 0.5 if len(vals) > 3 else 0.5)
        ax.set_xlim(0.45, len(vals) + 0.55)
        set_y_ticks(ax, spec, y_range(vals, spec), 10)
        add_limits(fig, ax, spec, W)
        ax.set_xticks(range(1, len(vals) + 1), labels)
    if spec.get("ylabel"):
        fig.text(10 / W, (bottom + (H - top - bottom) / 2) / H, spec["ylabel"], rotation=90,
                 ha="center", va="center", fontsize=9.8)
    fig.savefig(path, dpi=200)
    plt.close(fig)


# ============================== capability page ============================

def jmp_quantile(x, p):
    """JMP's definition: interpolate at position p*(n+1)."""
    x = np.sort(np.asarray(x, dtype=float))
    n = len(x)
    pos = p * (n + 1)
    if pos <= 1:
        return x[0]
    if pos >= n:
        return x[-1]
    lo = int(math.floor(pos))
    return x[lo - 1] + (pos - lo) * (x[lo] - x[lo - 1])


def nice_step(raw):
    k = 10 ** math.floor(math.log10(raw))
    for m in (1, 2, 2.5, 5, 10):
        if raw <= m * k * 1.0001:
            return m * k
    return 10 * k


def g(v, sig=8):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "."
    s = f"{v:.{sig}g}"
    if "e" in s:
        s = f"{v:.{sig}f}".rstrip("0").rstrip(".")
    return s


def capability_stats(x, side, k):
    x = np.asarray(x, dtype=float)
    n = len(x)
    mu, s = float(np.mean(x)), float(np.std(x, ddof=1))
    tq = stats.t.ppf(0.975, n - 1)
    lsl = mu - k * s if side in ("lower", "both") else None
    usl = mu + k * s if side in ("upper", "both") else None
    out = {
        "n": n, "mu": mu, "s": s, "sem": s / math.sqrt(n),
        "mu_lo": mu - tq * s / math.sqrt(n), "mu_hi": mu + tq * s / math.sqrt(n),
        "s_lo": math.sqrt((n - 1) * s * s / stats.chi2.ppf(0.975, n - 1)),
        "s_hi": math.sqrt((n - 1) * s * s / stats.chi2.ppf(0.025, n - 1)),
        "m2ll": n * math.log(2 * math.pi * s * s) + (n - 1),
        "lsl": lsl, "usl": usl,
        "act_lo": (x < lsl).mean() * 100 if lsl is not None else None,
        "act_hi": (x > usl).mean() * 100 if usl is not None else None,
        "cpl": (mu - lsl) / (3 * s) if lsl is not None else None,
        "cpu": (usl - mu) / (3 * s) if usl is not None else None,
        "exp_lo": stats.norm.cdf((lsl - mu) / s) * 100 if lsl is not None else None,
        "exp_hi": stats.norm.sf((usl - mu) / s) * 100 if usl is not None else None,
    }
    out["cp"] = (usl - lsl) / (6 * s) if (lsl is not None and usl is not None) else None
    out["cpk"] = min(v for v in (out["cpl"], out["cpu"]) if v is not None)
    acts = [v for v in (out["act_lo"], out["act_hi"]) if v is not None]
    out["act_tot"] = sum(acts)
    exps = [v for v in (out["exp_lo"], out["exp_hi"]) if v is not None]
    out["exp_tot"] = sum(exps)
    q = sorted([0.5, 0.0, 1.0, 0.995, 0.975, 0.9, 0.75, 0.25, 0.1, 0.025, 0.005])
    out["quant"] = {p: jmp_quantile(x, p) for p in q}
    return out


def sigma_quality(pct):
    if pct is None:
        return None
    p = pct / 100
    if p <= 0:
        return None
    return stats.norm.isf(p) + 1.5


def capability_png(x, name, side, k, xlabel, path, note=""):
    W, H = 1709, 468 + (15 if side == "both" else 0)   # JMP adds a row for two-sided specs
    fig = plt.figure(figsize=(W / 100, H / 100), dpi=100)
    st = capability_stats(x, side, k)
    x = np.asarray(x, dtype=float)

    def R(px, py, pw, ph, fc, ec=None, lw=0.6):
        fig.add_artist(Rectangle((px / W, 1 - (py + ph) / H), pw / W, ph / H, transform=fig.transFigure,
                                 fc=fc, ec=ec or "none", lw=lw))

    def T(px, py, s, size=8.4, ha="left", bold=False, color="black"):
        fig.text(px / W, 1 - py / H, s, ha=ha, va="center", fontsize=size, color=color,
                 weight="bold" if bold else "normal")

    def header(px, py, pw, label, size=9.2):
        R(px, py - 9, pw, 17, HDR_BG, HDR_EC)
        T(px + 4, py, label, size=size, bold=True)

    def axes(px, py, pw, ph):
        return fig.add_axes([px / W, 1 - (py + ph) / H, pw / W, ph / H])

    # title bar
    R(8, 4, W - 12, 18, "#F2F2F2", "#C8C8C8")
    T(14, 13, f"{name}-{k}sigma", size=10, bold=True)
    if note:
        T(W - 12, 13, note, size=8, ha="right", color="#707070")

    # ---- histogram + outlier box plot ----
    limits = [v for v in (st["lsl"], st["usl"]) if v is not None]
    lo, hi = min(x.min(), *limits), max(x.max(), *limits)
    target = min(25, max(6, 1 + 2 * math.log2(len(x))))
    bw = nice_step((x.max() - x.min()) / target if x.max() > x.min() else 1)
    start = math.floor(x.min() / bw) * bw
    edges = np.arange(start, x.max() + bw * 1.0001, bw)
    if len(edges) < 2:
        edges = np.array([start, start + bw])
    xl = (min(lo, edges[0]) - bw * 0.6, max(hi, edges[-1]) + bw * 0.6)

    bx = axes(30, 32, 277, 45)
    hx = axes(30, 77, 277, 115)
    for a in (bx, hx):
        for sp in a.spines.values():
            sp.set_color(FRAME)
            sp.set_linewidth(0.7)
        a.set_xlim(*xl)
        a.set_yticks([])
    bx.set_xticks([])
    bx.set_ylim(0, 1)
    q1, med, q3 = (jmp_quantile(x, p) for p in (0.25, 0.5, 0.75))
    iqr = q3 - q1
    inside = x[(x >= q1 - 1.5 * iqr) & (x <= q3 + 1.5 * iqr)]
    wl, wh = inside.min(), inside.max()
    yc = 0.45
    bx.plot([wl, q1], [yc, yc], color="black", lw=0.8)
    bx.plot([q3, wh], [yc, yc], color="black", lw=0.8)
    for e in (wl, wh):
        bx.plot([e, e], [yc - 0.12, yc + 0.12], color="black", lw=0.8)
    bx.add_patch(Rectangle((q1, yc - 0.17), q3 - q1, 0.34, fc="white", ec="black", lw=0.8))
    bx.plot([med, med], [yc - 0.17, yc + 0.17], color="black", lw=0.8)
    dh = st["mu_hi"] - st["mu"]
    bx.add_patch(Polygon([[st["mu"] - dh, yc], [st["mu"], yc + 0.13], [st["mu"] + dh, yc],
                          [st["mu"], yc - 0.13]], closed=True, fc="none", ec="black", lw=0.7))
    xs = np.sort(x)
    h = max(int(math.ceil(len(xs) / 2)), 1)
    if len(xs) > h:
        widths = xs[h - 1:] - xs[:len(xs) - h + 1]
        j = int(np.argmin(widths))
        a_, b_ = xs[j], xs[j + h - 1]
        bx.plot([a_, b_], [0.82, 0.82], color=RED, lw=0.8)
        bx.plot([a_, a_], [0.82, 0.74], color=RED, lw=0.8)
        bx.plot([b_, b_], [0.82, 0.74], color=RED, lw=0.8)
    outs = x[(x < wl) | (x > wh)]
    bx.plot(outs, np.full(len(outs), yc), "o", ms=2.8, color="black")

    counts, _, _ = hx.hist(x, bins=edges, color=BAR, edgecolor=BAR_EDGE, linewidth=0.5)
    gx = np.linspace(*xl, 300)
    gy = stats.norm.pdf(gx, st["mu"], st["s"]) * len(x) * bw
    hx.plot(gx, gy, color=RED, lw=1.4)
    hx.set_ylim(0, max(counts.max(), gy.max()) * 1.12)
    for v, lab in ((st["lsl"], "LSL"), (st["usl"], "USL")):
        if v is not None:
            hx.axvline(v, color=BLUE, lw=0.8)
            hx.text(v, 0.86, lab, transform=hx.get_xaxis_transform(), ha="center", va="center",
                    fontsize=6.3, bbox={"fc": "white", "ec": "none", "pad": 0.2})
    hx.xaxis.set_major_locator(MaxNLocator(nbins=8, steps=[1, 2, 2.5, 5, 10]))
    hx.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:g}"))
    hx.tick_params(axis="x", length=3, width=0.6, labelsize=8.2, pad=2, color="#606060")
    fig.add_artist(plt.Line2D([14 / W, 36 / W], [1 - 245 / H, 1 - 245 / H], color=RED, lw=1,
                              transform=fig.transFigure))
    T(40, 245, f"Normal({st['mu']:.6g},{st['s']:.6g})", size=8.2)

    # ---- quantiles ----
    header(337, 36, 193, "Quantiles")
    R(337, 47, 46, 166, LABEL_BG)
    rows = [(1.0, "100.0%", "maximum"), (0.995, "99.5%", ""), (0.975, "97.5%", ""), (0.9, "90.0%", ""),
            (0.75, "75.0%", "quartile"), (0.5, "50.0%", "median"), (0.25, "25.0%", "quartile"),
            (0.1, "10.0%", ""), (0.025, "2.5%", ""), (0.005, "0.5%", ""), (0.0, "0.0%", "minimum")]
    for i, (p, pl, lab) in enumerate(rows):
        y = 55 + 15 * i
        T(339, y, pl)
        if lab:
            T(452, y, lab, ha="right")
        T(525, y, g(st["quant"][p], 7), ha="right")

    # ---- summary ----
    header(540, 36, 168, "Summary Statistics")
    R(540, 47, 92, 91, LABEL_BG)
    for i, (lab, v) in enumerate([("Mean", g(st["mu"])), ("Std Dev", g(st["s"])), ("Std Err Mean", g(st["sem"])),
                                  ("Upper 95% Mean", g(st["mu_hi"])), ("Lower 95% Mean", g(st["mu_lo"])),
                                  ("N", str(st["n"]))]):
        y = 55 + 15 * i
        T(542, y, lab, size=7.9)
        T(704, y, v, ha="right")

    # ---- CDF ----
    header(716, 36, 391, "CDF Plot")
    cx = axes(770, 55, 323, 240)
    for sp in cx.spines.values():
        sp.set_color(FRAME)
        sp.set_linewidth(0.7)
    cy = np.arange(1, len(xs) + 1) / len(xs)
    cx.step(np.r_[xs[0] - (xs[-1] - xs[0]) * 0.03, xs], np.r_[0, cy], where="post", color=RED, lw=1)
    cx.set_ylim(0, 1.02)
    cx.set_yticks([0.2, 0.4, 0.6, 0.8])
    cx.xaxis.set_major_locator(MaxNLocator(nbins=8, steps=[1, 2, 2.5, 5, 10]))
    cx.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:g}"))
    cx.tick_params(labelsize=8.2, length=3, width=0.6, color="#606060", pad=2)
    cx.set_xlabel(xlabel, fontsize=8.4, labelpad=2)
    T(728, 175, "")
    fig.text(728 / W, 1 - 175 / H, "Cum Prob", rotation=90, ha="center", va="center", fontsize=7.6)

    # ---- normal fit ----
    header(1120, 36, 585, "Fitted Normal")
    header(1132, 58, 335, "Parameter Estimates")
    R(1132, 70, 335, 15, HDR_BG)
    for px, s_, ha in ((1134, "Type", "left"), (1200, "Parameter", "left"), (1322, "Estimate", "right"),
                       (1393, "Lower 95%", "right"), (1463, "Upper 95%", "right")):
        T(px, 78, s_, size=8.2, bold=True, ha=ha)
    R(1132, 87, 70, 30, LABEL_BG)
    for y, a_, b_, c_, d_, e_ in ((95, "Location", "μ", st["mu"], st["mu_lo"], st["mu_hi"]),
                                  (110, "Dispersion", "σ", st["s"], st["s_lo"], st["s_hi"])):
        T(1134, y, a_)
        T(1200, y, b_, color="#B03030")
        T(1322, y, g(c_), ha="right")
        T(1393, y, g(d_), ha="right")
        T(1463, y, g(e_), ha="right")
    T(1132, 127, f"-2log(Likelihood) = {st['m2ll']:.15g}")

    header(1132, 147, 573, "Process Capability")
    R(1132, 159, 563, 15, HDR_BG)
    for px, s_, ha in ((1134, "Specification", "left"), (1290, "Value", "right"), (1300, "Portion", "left"),
                       (1450, "% Actual", "right")):
        T(px, 167, s_, size=8.2, bold=True, ha=ha)
    R(1132, 176, 98, 45, LABEL_BG)
    act = lambda v: "." if v is None else f"{v:.4f}"
    for y, a_, v, b_, c_ in ((184, "Lower Spec Limit", st["lsl"], "Below LSL", act(st["act_lo"])),
                             (199, "Spec Target", None, "Above USL", act(st["act_hi"])),
                             (214, "Upper Spec Limit", st["usl"], "Total Outside", f"{st['act_tot']:.4f}")):
        T(1134, y, a_)
        T(1290, y, "." if v is None else g(v, 7), ha="right")
        T(1300, y, b_)
        T(1450, y, c_, ha="right")

    header(1145, 234, 560, "Quantile Sigma", size=9.6)
    sx = axes(1155, 250, 232, 140)
    sx.axis("off")
    mu, s = st["mu"], st["s"]
    lo_s = min(mu - 4.2 * s, *(limits or [mu]))
    hi_s = max(mu + 4.2 * s, *(limits or [mu]))
    span = hi_s - lo_s
    sx.set_xlim(lo_s - span * 0.03, hi_s + span * 0.03)
    sx.set_ylim(-0.75, 1.15)
    zz = np.linspace(mu - 4 * s, mu + 4 * s, 200)
    sx.plot(zz, np.exp(-0.5 * ((zz - mu) / s) ** 2), color="black", lw=0.8)
    sx.plot([lo_s - span * 0.03, hi_s + span * 0.03], [0, 0], color="black", lw=0.6)
    sx.plot([mu - 3 * s, mu + 3 * s], [0.0, 0.0], color="black", lw=0.8)
    for v, lab in ((mu - 3 * s, "-3s"), (mu, "Mean"), (mu + 3 * s, "+3s")):
        sx.plot([v, v], [-0.04, 0.12], color="black", lw=0.8)
        sx.text(v, 0.2, lab, ha="center", va="center", fontsize=6.3)
    for v, lab in ((st["lsl"], "LSL"), (st["usl"], "USL")):
        if v is not None:
            sx.plot([v, v], [0.12, -0.25], color="black", lw=0.8)
            sx.text(v, -0.35, lab, ha="center", va="center", fontsize=8)
    sx.plot([lo_s - span * 0.03, hi_s + span * 0.03], [-0.48, -0.48], color="black", lw=0.6)
    tick_loc = MaxNLocator(nbins=3, steps=[1, 2, 2.5, 5, 10]).tick_values(lo_s, hi_s)
    for t in tick_loc:
        if lo_s <= t <= hi_s:
            sx.plot([t, t], [-0.48, -0.53], color="black", lw=0.6)
            sx.text(t, -0.66, f"{t:g}", ha="center", va="center", fontsize=8.2)

    R(1415, 246, 108, 15, HDR_BG)
    T(1417, 254, "Capability", size=8.2, bold=True)
    T(1518, 254, "Index", size=8.2, bold=True, ha="right")
    R(1415, 263, 30, 75, LABEL_BG)
    idx = lambda v: "." if v is None else f"{v:.3f}"
    for i, (lab, v) in enumerate((("CP", st["cp"]), ("CPK", st["cpk"]), ("CPM", None),
                                  ("CPL", st["cpl"]), ("CPU", st["cpu"]))):
        T(1417, 271 + 15 * i, lab)
        T(1518, 271 + 15 * i, idx(v), ha="right")
    R(1415, 340, 290, 15, HDR_BG)
    for px, s_, ha in ((1417, "Portion", "left"), (1540, "Percent", "right"), (1606, "PPM", "right"),
                       (1702, "Sigma Quality", "right")):
        T(px, 348, s_, size=8.2, bold=True, ha=ha)
    R(1415, 357, 78, 45, LABEL_BG)
    pc = lambda v: "." if v is None else f"{v:.4f}"
    sq = lambda v: "." if v is None else f"{sigma_quality(v):.3f}"
    for i, (lab, v) in enumerate((("Below LSL", st["exp_lo"]), ("Above USL", st["exp_hi"]),
                                  ("Total Outside", st["exp_tot"]))):
        y = 365 + 15 * i
        T(1417, y, lab)
        T(1540, y, pc(v), ha="right")
        T(1606, y, "." if v is None else f"{v * 1e4:.4f}", ha="right")
        T(1702, y, sq(v), ha="right")

    header(1132, 416, 235, "Quantiles: Uncentered and Unscaled")
    T(1195, 437, "Percentile", size=8.2, bold=True, ha="right")
    T(1265, 437, "Quantile", size=8.2, bold=True, ha="right")
    rows = []
    if st["lsl"] is not None:
        rows.append((stats.norm.cdf(-k * 1.0), st["lsl"]))
    if st["usl"] is not None:
        rows.append((stats.norm.cdf(k * 1.0), st["usl"]))
    for i, (p, v) in enumerate(rows):
        T(1195, 453 + 15 * i, f"{p:.7f}", ha="right")
        T(1265, 453 + 15 * i, g(v, 7), ha="right")

    fig.savefig(path, dpi=150)
    plt.close(fig)


def capability_values(df, spec_key, cs, cfg):
    """Values for a capability fit. Failed runs are left out by default: an
    aborted or unpowered measurement (e.g. 7.9 W when good units draw ~24 W)
    is not part of the population and would wreck sigma."""
    spec = cfg["charts"][spec_key]
    excluded = 0
    if cfg.get("capability_exclude_failed", True) and "_pass" in df:
        cols, _ = resolve_columns(df, spec, cfg)
        if cols:
            has = df[[c for _, c in cols]].apply(pd.to_numeric, errors="coerce").notna().any(axis=1)
            scope = df["_temp"] > 50 if (cs.get("temp") == "high" and (df["_temp"] > 50).any()) else df["_temp"].notna()
            excluded = int((has & scope & ~df["_pass"]).sum())
        df = df[df["_pass"]]
    cd = chart_data(df, spec, cfg)
    panels = cd["panels"]
    if cs.get("temp") == "high" and spec["kind"] != "power_drop":
        hi = [p for p in panels if p[0] is not None and p[0] > 50]
        prefix = "HT " if hi else ""
        panels = hi or panels
    else:
        prefix = ""
    vals = pd.concat([pd.Series(v, dtype=float) for _, _, lanes in panels for v in lanes]) if panels else pd.Series(dtype=float)
    pooled = spec["kind"] in ("lanes", "power_drop")
    xlabel = "Data" if pooled else to_english(cd["info"].get("base", spec_key))
    if cs["chart"] == "TDSP":
        xlabel = "TDSP"
    return vals.dropna().values, prefix, xlabel, cd, excluded


# ============================== deck =======================================

FONT_EN = {"ＭＳ Ｐゴシック": "MS PGothic", "宋体": "SimSun", "新細明體": "PMingLiU", "微软雅黑": "Microsoft YaHei",
           "等线": "DengXian", "黑体": "SimHei", "맑은 고딕": "Malgun Gothic"}


def english_package(path):
    """python-pptx's default theme names East-Asian fallback fonts in CJK
    script. Rewrite them to their English family names so the file is
    English end to end."""
    import shutil
    import zipfile
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml"):
                txt = data.decode("utf-8")
                for zh, en in FONT_EN.items():
                    txt = txt.replace(f'typeface="{zh}"', f'typeface="{en}"')
                txt = re.sub(r'<a:font script="\w+" typeface="[^"]*[\u3000-\u9fff\uac00-\ud7af\uff00-\uffef][^"]*"/>', "", txt)
                data = txt.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)


def build_deck(df, cfg, out_path, template):
    prs = Presentation(template)
    title_slide, proto = prs.slides[0], prs.slides[1]
    proto_title = copy.deepcopy(next(sh for sh in proto.shapes if sh.has_text_frame)._element)
    layout = proto.slide_layout

    # title slide: first run = product, last run = " Data Distribution"
    for sh in title_slide.shapes:
        if sh.has_text_frame and "Distribution" in sh.text_frame.text:
            runs = sh.text_frame.paragraphs[0].runs
            for r in runs:
                r.text = ""
            runs[0].text = cfg["product"]
            if len(runs) > 1:
                runs[-1].text = " Data Distribution"
            else:
                runs[0].text += " Data Distribution"

    def new_slide(title):
        s = prs.slides.add_slide(layout)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        el = copy.deepcopy(proto_title)
        s.shapes._spTree.append(el)
        tb = s.shapes[-1]
        tb.text_frame.paragraphs[0].runs[0].text = title
        return s

    notes = []
    tmp = tempfile.mkdtemp(prefix="dr8_")
    for si, ov in enumerate(cfg["overview_slides"]):
        s = new_slide(ov["title"])
        for ci, (key, x, y) in enumerate(ov["charts"]):
            spec = cfg["charts"][key]
            cd = chart_data(df, spec, cfg)
            if cd["info"].get("note"):
                notes.append(f"{cd['title']}: {cd['info']['note']}")
            png = os.path.join(tmp, f"ov{si}_{ci}.png")
            if spec["kind"] == "lanes":
                variability_png(cd, spec, ov["w"], ov["h"], png)
            else:
                single_png(cd, spec, ov["w"], ov["h"], png)
            s.shapes.add_picture(png, Inches(x), Inches(y), Inches(ov["w"]), Inches(ov["h"]))

    cp = cfg["capability_placement"]
    for i, cs in enumerate(cfg["capability_slides"]):
        vals, prefix, xlabel, cd, excluded = capability_values(df, cs["chart"], cs, cfg)
        note = f"Fit excludes {excluded} failed (NG) run{'s' if excluded != 1 else ''}" if excluded else ""
        if excluded:
            notes.append(f"{cs['slide_title']}: {note[0].lower() + note[1:]}")
        s = new_slide(cs["slide_title"])
        if len(vals) < 3 or np.std(vals) == 0:
            notes.append(f"{cs['slide_title']}: not enough spread for a capability fit, page left blank")
            continue
        for j, k in enumerate(cfg["sigma_levels"][:2]):
            png = os.path.join(tmp, f"cap{i}_{k}.png")
            capability_png(vals, f"{prefix}{cs['name']}", cs["side"], k, xlabel, png, note)
            y = cp["y_top"] if j == 0 else cp["y_bottom"]
            h = cp["w"] * (483 if cs["side"] == "both" else 468) / 1709
            s.shapes.add_picture(png, Inches(cp["x"]), Inches(y), Inches(cp["w"]), Inches(h))

    # drop the prototype slide that only carried the title styling
    sld_ids = prs.slides._sldIdLst
    sid = list(sld_ids)[1]
    prs.part.drop_rel(sid.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"))
    sld_ids.remove(sid)
    prs.save(out_path)
    english_package(out_path)
    return notes


# ============================== CLI ========================================

def load_config(path):
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    if path:
        with open(path, encoding="utf-8") as f:
            user = json.load(f)
        for k, v in user.items():
            if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                cfg[k].update(v)
            else:
                cfg[k] = v
    return cfg


def ask(prompt, default=""):
    try:
        v = input(f"{prompt}{f' [{default}]' if default else ''}: ").strip()
    except EOFError:
        v = ""
    return v or default


def pick(label, counts, preselect):
    items = list(counts.items())
    print(f"\n{label} found in the data:")
    for i, (k, n) in enumerate(items, 1):
        mark = "*" if k in preselect else " "
        print(f"  {mark}{i:>2}. {k}  ({n} runs)")
    raw = ask(f"Pick {label.lower()} (numbers/names, comma separated; blank = the * ones, 'all' = everything)")
    if not raw:
        return preselect or None
    if raw.lower() == "all":
        return None
    chosen = []
    for tok in [t.strip() for t in raw.split(",") if t.strip()]:
        if tok.isdigit() and 1 <= int(tok) <= len(items):
            chosen.append(items[int(tok) - 1][0])
        elif tok in counts:
            chosen.append(tok)
    return chosen or None


def main():
    ap = argparse.ArgumentParser(description="Build the JMP-style DR8 distribution deck from test-data exports.")
    ap.add_argument("files", nargs="*", help="Excel/CSV test-data export(s)")
    ap.add_argument("-o", "--output", help="output .pptx path")
    ap.add_argument("--product", help="product name for the title slide (default DR8)")
    ap.add_argument("--config", help="JSON config with limits, charts, capability pages")
    ap.add_argument("--dump-config", metavar="PATH", help="write the default config and exit")
    ap.add_argument("--template", help="template .pptx (default dr8_template.pptx next to this script)")
    ap.add_argument("--process", action="append", help="keep only this process (repeatable)")
    ap.add_argument("--all-processes", action="store_true", help="ignore the default final-test scope")
    ap.add_argument("--from", dest="date_from", help="start date YYYY-MM-DD")
    ap.add_argument("--to", dest="date_to", help="end date YYYY-MM-DD")
    ap.add_argument("--no-prompt", action="store_true", help="never ask questions")
    a = ap.parse_args()

    if a.dump_config:
        with open(a.dump_config, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, indent=2, ensure_ascii=False)
        print(f"Default config written to {a.dump_config}")
        return

    interactive = not a.no_prompt and sys.stdin.isatty()
    files = a.files
    if not files and interactive:
        files = [p.strip().strip('"') for p in ask("Path to test-data Excel/CSV (comma separate several)").split(",")
                 if p.strip()]
    if not files:
        sys.exit("No input file given.")
    for p in files:
        if not os.path.exists(p):
            sys.exit(f"File not found: {p}")

    cfg = load_config(a.config or (ask("Config JSON (blank = built-in limits)") if interactive else None))
    if a.product:
        cfg["product"] = a.product
    elif interactive:
        cfg["product"] = ask("Product name for the title slide", cfg["product"])
    template = a.template or cfg["template"]
    if not os.path.isabs(template) and not os.path.exists(template):
        template = os.path.join(HERE, template)
    if not os.path.exists(template):
        sys.exit(f"Template not found: {template}")

    df = prep(load_frames(files), cfg)
    procs = a.process
    if not procs and not a.all_processes:
        rx = re.compile(cfg["scope_process_regex"], re.I) if cfg.get("scope_process_regex") else None
        default = [p for p in df["_process"].unique() if rx and rx.search(p)]
        if interactive:
            procs = pick("Processes", df["_process"].value_counts().to_dict(), default)
        else:
            procs = default or None
    if procs:
        df = df[df["_process"].isin(procs)]
    date_from, date_to = a.date_from, a.date_to
    if interactive and not (date_from or date_to):
        date_from = ask("Start date YYYY-MM-DD (blank = all)") or None
        date_to = ask("End date YYYY-MM-DD (blank = all)") or None
    if date_from:
        df = df[df["_start"] >= pd.to_datetime(date_from)]
    if date_to:
        df = df[df["_start"] < pd.to_datetime(date_to) + pd.Timedelta(days=1)]
    if df.empty:
        sys.exit("Filters removed every row.")
    df = df.reset_index(drop=True)

    out = a.output or (ask("Output file", f"{cfg['product']}_Data_Distribution.pptx") if interactive
                       else f"{cfg['product']}_Data_Distribution.pptx")
    if not out.lower().endswith(".pptx"):
        out += ".pptx"
    notes = build_deck(df, cfg, out, template)
    temps = ", ".join(f"{t:g}°C" for t in temp_groups(df))
    print(f"\nSaved {out}")
    print(f"  {len(df)} runs, {df['_serial'].nunique()} units, processes: {', '.join(sorted(df['_process'].unique()))}")
    print(f"  temperatures: {temps or 'none found'}")
    for n in notes:
        print(f"  note: {n}")


if __name__ == "__main__":
    main()
