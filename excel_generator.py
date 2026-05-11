"""
excel_generator.py  –  v1.0.1
Genererer nettoomsetningsrapporten med 7 regneark.
Bruker dict fra data_processor.process().

Litteratur og rammeverk:
  - ABC-analyse:        Dickie (1951) — selektiv lagerstyring
  - HHI / CR3/CR5:     Herfindahl-Hirschman (1964); Utton (1975)
  - Pareto 80/20:       Juran & Godfrey (1999)
  - Gini-koeffisient:   Gini (1912)
  - XYZ-analyse:        Silver, Pyke & Peterson (1998); Scholz-Reiter et al. (2012)
  - BCG-matrise:        Henderson (1970)
  - CAGR:               Standard finansiell vekstbenchmarking
  - Sesongindeks:       Makridakis, Wheelwright & Hyndman (1998)
"""
import io
import math
from datetime import datetime

import pandas as pd
import xlsxwriter

APP_VERSION = "1.0.1"

# ── Måneder ──────────────────────────────────────────────────────────────────
MONTHS_EN = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
MONTHS_NO = ["Jan","Feb","Mar","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Des"]
_M_MAP = dict(zip(MONTHS_EN, MONTHS_NO))

# ── Fargepalett ───────────────────────────────────────────────────────────────
DARK_BLUE    = "#1A3A5C"
MID_BLUE     = "#2E75B6"
LIGHT_BLUE   = "#D6E4F0"
ACCENT_GREEN = "#1E8449"
ACCENT_RED   = "#C0392B"
ACCENT_ORG   = "#D35400"
WHITE        = "#FFFFFF"
LIGHT_GREY   = "#F2F2F2"
MID_GREY     = "#D9D9D9"
DARK_GREY    = "#595959"
GOLD_BG      = "#FFF2CC"
GOLD_BORDER  = "#F1C40F"
SEAS_GREEN   = "#C6EFCE"
SEAS_RED     = "#FFC7CE"
PURPLE       = "#7B2D8B"
TEAL         = "#1A7A72"

# BCG-kategorifarger
_BCG_COLORS = {
    "Stjerne":       "#F4D03F",   # gull
    "Melkeku":       "#1E8449",   # grønn
    "Spørsmålstegn": "#2E75B6",   # blå
    "Hund":          "#7F8C8D",   # grå
    "Ny":            "#D35400",   # oransje
}
_BCG_MARKERS = {
    "Stjerne":       "diamond",
    "Melkeku":       "circle",
    "Spørsmålstegn": "square",
    "Hund":          "x",
    "Ny":            "triangle",
}


# ── Hjelpefunksjoner ──────────────────────────────────────────────────────────

def _v(val):
    if val is None:
        return ""
    try:
        if math.isnan(val):
            return ""
    except (TypeError, ValueError):
        pass
    return val


def _pct(val):
    return _v(val)


def _nok(v):
    try:
        return f"NOK {v:,.0f}"
    except Exception:
        return str(v)


def _add_formats(wb):
    f = {}

    # ── Tittel / banner ──────────────────────────────────────────────────
    f["title"] = wb.add_format({
        "bold": True, "font_size": 16, "font_name": "Calibri",
        "font_color": WHITE, "bg_color": DARK_BLUE,
        "align": "left", "valign": "vcenter",
    })
    f["subtitle"] = wb.add_format({
        "font_size": 9, "font_name": "Calibri",
        "font_color": "#AACCE8", "italic": True,
        "bg_color": DARK_BLUE, "align": "left", "valign": "vcenter",
    })
    f["section_hdr"] = wb.add_format({
        "bold": True, "font_size": 10, "font_name": "Calibri",
        "font_color": WHITE, "bg_color": MID_BLUE,
        "align": "left", "valign": "vcenter",
        "left": 3, "left_color": GOLD_BORDER,
    })
    f["section_hdr_green"] = wb.add_format({
        "bold": True, "font_size": 10, "font_name": "Calibri",
        "font_color": WHITE, "bg_color": ACCENT_GREEN,
        "align": "left", "valign": "vcenter",
        "left": 3, "left_color": GOLD_BORDER,
    })
    f["section_hdr_gold"] = wb.add_format({
        "bold": True, "font_size": 10, "font_name": "Calibri",
        "font_color": DARK_BLUE, "bg_color": GOLD_BG,
        "align": "left", "valign": "vcenter",
        "left": 3, "left_color": GOLD_BORDER,
        "border": 1, "border_color": MID_GREY,
    })

    # ── KPI-fliser ───────────────────────────────────────────────────────
    f["kpi_label"] = wb.add_format({
        "bold": True, "font_size": 9, "font_name": "Calibri",
        "font_color": DARK_GREY, "bg_color": LIGHT_BLUE,
        "align": "center", "valign": "bottom",
        "top": 2, "top_color": MID_BLUE,
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_value"] = wb.add_format({
        "bold": True, "font_size": 18, "font_name": "Calibri",
        "font_color": DARK_BLUE, "bg_color": LIGHT_BLUE,
        "align": "center", "valign": "vcenter",
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_value_nok"] = wb.add_format({
        "bold": True, "font_size": 14, "font_name": "Calibri",
        "font_color": DARK_BLUE, "bg_color": LIGHT_BLUE,
        "align": "center", "valign": "vcenter",
        "num_format": '#,##0 "NOK"',
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_value_pct"] = wb.add_format({
        "bold": True, "font_size": 18, "font_name": "Calibri",
        "font_color": DARK_BLUE, "bg_color": LIGHT_BLUE,
        "align": "center", "valign": "vcenter",
        "num_format": "0.0%",
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_text"] = wb.add_format({
        "bold": True, "font_size": 11, "font_name": "Calibri",
        "font_color": DARK_BLUE, "bg_color": LIGHT_BLUE,
        "align": "center", "valign": "vcenter",
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_delta_pos"] = wb.add_format({
        "font_size": 9, "font_name": "Calibri",
        "font_color": ACCENT_GREEN, "bg_color": LIGHT_BLUE,
        "align": "center", "valign": "top",
        "num_format": '+0.0%;-0.0%',
        "bottom": 2, "bottom_color": MID_BLUE,
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_delta_neg"] = wb.add_format({
        "font_size": 9, "font_name": "Calibri",
        "font_color": ACCENT_RED, "bg_color": LIGHT_BLUE,
        "align": "center", "valign": "top",
        "num_format": '+0.0%;-0.0%',
        "bottom": 2, "bottom_color": MID_BLUE,
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_delta_na"] = wb.add_format({
        "font_size": 9, "font_name": "Calibri",
        "font_color": DARK_GREY, "bg_color": LIGHT_BLUE,
        "align": "center", "valign": "top",
        "bottom": 2, "bottom_color": MID_BLUE,
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })

    # ── Sparkline-rad ────────────────────────────────────────────────────
    f["sparkline_label"] = wb.add_format({
        "font_size": 8, "font_name": "Calibri", "italic": True,
        "font_color": DARK_GREY, "bg_color": "#EBF3FB",
        "align": "center", "valign": "vcenter",
        "bottom": 2, "bottom_color": MID_BLUE,
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["sparkline_cell"] = wb.add_format({
        "bg_color": "#EBF3FB",
        "bottom": 2, "bottom_color": MID_BLUE,
        "right": 1, "right_color": MID_BLUE,
    })

    # ── Kolonneoverskrifter ──────────────────────────────────────────────
    f["col_hdr"] = wb.add_format({
        "bold": True, "font_size": 10, "font_name": "Calibri",
        "font_color": WHITE, "bg_color": DARK_BLUE,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": MID_BLUE, "text_wrap": True,
    })
    f["col_hdr_left"] = wb.add_format({
        "bold": True, "font_size": 10, "font_name": "Calibri",
        "font_color": WHITE, "bg_color": DARK_BLUE,
        "align": "left", "valign": "vcenter",
        "border": 1, "border_color": MID_BLUE,
    })
    f["col_hdr_yr"] = wb.add_format({
        "bold": True, "font_size": 9, "font_name": "Calibri",
        "font_color": WHITE, "bg_color": MID_BLUE,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": MID_BLUE, "text_wrap": True,
    })

    # ── Tabellceller ─────────────────────────────────────────────────────
    def _cell(bold=False, align="left", num_format=None, bg=None, wrap=False, font_name="Calibri"):
        d = {
            "font_size": 10, "font_name": font_name,
            "align": align, "valign": "vcenter",
            "border": 1, "border_color": MID_GREY,
        }
        if bold:       d["bold"] = True
        if num_format: d["num_format"] = num_format
        if bg:         d["bg_color"] = bg
        if wrap:       d["text_wrap"] = True
        return wb.add_format(d)

    f["cell"]           = _cell()
    f["cell_c"]         = _cell(align="center")
    f["cell_nok"]       = _cell(align="right",  num_format='#,##0 "NOK"')
    f["cell_pct"]       = _cell(align="center", num_format="0.0%")
    f["cell_pct_yoy"]   = _cell(align="center", num_format='+0.0%;-0.0%;"-"')
    f["cell_int"]       = _cell(align="center", num_format="#,##0")
    f["cell_2dec"]      = _cell(align="center", num_format="0.00")
    f["cell_wrap"]      = _cell(wrap=True)

    f["cell_a"]         = _cell(bg=LIGHT_GREY)
    f["cell_c_a"]       = _cell(align="center", bg=LIGHT_GREY)
    f["cell_nok_a"]     = _cell(align="right",  num_format='#,##0 "NOK"', bg=LIGHT_GREY)
    f["cell_pct_a"]     = _cell(align="center", num_format="0.0%",        bg=LIGHT_GREY)
    f["cell_pct_yoy_a"] = _cell(align="center", num_format='+0.0%;-0.0%;"-"', bg=LIGHT_GREY)
    f["cell_int_a"]     = _cell(align="center", num_format="#,##0",       bg=LIGHT_GREY)
    f["cell_2dec_a"]    = _cell(align="center", num_format="0.00",        bg=LIGHT_GREY)

    f["cell_empty_yr"]  = _cell(align="center", bg="#F9F9F9")

    # Totalrader
    f["total"]          = _cell(bold=True, bg=LIGHT_BLUE)
    f["total_c"]        = _cell(bold=True, align="center", bg=LIGHT_BLUE)
    f["total_nok"]      = _cell(bold=True, align="right",  num_format='#,##0 "NOK"', bg=LIGHT_BLUE)
    f["total_pct"]      = _cell(bold=True, align="center", num_format="0.0%",        bg=LIGHT_BLUE)
    f["total_pct_yoy"]  = _cell(bold=True, align="center", num_format='+0.0%;-0.0%;"-"', bg=LIGHT_BLUE)
    f["total_int"]      = _cell(bold=True, align="center", num_format="#,##0",       bg=LIGHT_BLUE)

    # Pareto 80%-rad
    f["p80_c"]   = _cell(bold=True, align="center", bg=GOLD_BG)
    f["p80_l"]   = _cell(bold=True, bg=GOLD_BG)
    f["p80_nok"] = _cell(bold=True, align="right",  num_format='#,##0 "NOK"', bg=GOLD_BG)
    f["p80_pct"] = _cell(bold=True, align="center", num_format="0.0%",        bg=GOLD_BG)
    f["p80_tag"] = wb.add_format({
        "bold": True, "font_size": 9, "font_color": "#7B241C",
        "bg_color": GOLD_BG, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": MID_GREY,
    })

    # Sesongindeks
    f["seas_hi"]  = _cell(align="center", num_format="0.00", bg=SEAS_GREEN)
    f["seas_lo"]  = _cell(align="center", num_format="0.00", bg=SEAS_RED)
    f["seas_mid"] = _cell(align="center", num_format="0.00")

    # ABC-merker
    for k, bg, fc in [("A", ACCENT_GREEN, WHITE), ("B", MID_BLUE, WHITE), ("C", DARK_GREY, WHITE)]:
        f[f"abc_{k}"] = wb.add_format({"bold": True, "font_size": 10, "font_name": "Calibri",
                                        "align": "center", "valign": "vcenter",
                                        "font_color": fc, "bg_color": bg,
                                        "border": 1, "border_color": MID_GREY})

    # XYZ-merker
    for k, bg, fc in [("X", TEAL, WHITE), ("Y", GOLD_BG, DARK_BLUE), ("Z", ACCENT_RED, WHITE)]:
        f[f"xyz_{k}"] = wb.add_format({"bold": True, "font_size": 10, "font_name": "Calibri",
                                        "align": "center", "valign": "vcenter",
                                        "font_color": fc, "bg_color": bg,
                                        "border": 1, "border_color": MID_GREY})

    # Portfolio-kategorier
    cat_defs = [
        ("cat_stjerne",  GOLD_BG,      DARK_BLUE),
        ("cat_melkeku",  ACCENT_GREEN,  WHITE),
        ("cat_spm",      MID_BLUE,      WHITE),
        ("cat_hund",     DARK_GREY,     WHITE),
        ("cat_ny",       ACCENT_ORG,    WHITE),
    ]
    for key, bg, fc in cat_defs:
        f[key] = wb.add_format({"bold": True, "font_size": 10, "font_name": "Calibri",
                                 "align": "center", "valign": "vcenter",
                                 "font_color": fc, "bg_color": bg,
                                 "border": 1, "border_color": MID_GREY})

    # Rangmerker
    f["rank_gold"] = wb.add_format({
        "bold": True, "font_size": 10, "font_name": "Calibri",
        "font_color": DARK_BLUE, "bg_color": GOLD_BG,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": MID_GREY,
    })
    f["rank_std"] = wb.add_format({
        "bold": True, "font_size": 10, "font_name": "Calibri",
        "font_color": WHITE, "bg_color": MID_BLUE,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": MID_GREY,
    })

    f["bullet"] = wb.add_format({
        "font_size": 10, "font_name": "Calibri",
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    f["blank"]        = wb.add_format({"bg_color": WHITE})
    f["blank_dark"]   = wb.add_format({"bg_color": DARK_BLUE})
    f["note"]         = wb.add_format({
        "italic": True, "font_size": 8, "font_name": "Calibri",
        "font_color": DARK_GREY, "align": "left", "valign": "vcenter",
        "indent": 1, "text_wrap": True,
    })
    f["insight_bold"] = wb.add_format({
        "bold": True, "font_size": 10, "font_name": "Calibri",
        "align": "left", "valign": "vcenter", "indent": 1,
        "font_color": WHITE, "bg_color": DARK_BLUE,
    })
    f["insight_body"] = wb.add_format({
        "font_size": 10, "font_name": "Calibri",
        "align": "left", "valign": "vcenter", "indent": 2,
        "font_color": "#1A1A2E", "text_wrap": True,
        "bg_color": LIGHT_BLUE,
        "border": 1, "border_color": MID_GREY,
    })

    return f


def _write_kpi_tile(ws, fmt, row, col, label, value, val_fmt_key, delta=None):
    ws.merge_range(row,   col, row,   col+1, label, fmt["kpi_label"])
    ws.merge_range(row+1, col, row+1, col+1, value, fmt[val_fmt_key])
    if delta is None:
        ws.merge_range(row+2, col, row+2, col+1, "", fmt["kpi_delta_na"])
    elif delta >= 0:
        ws.merge_range(row+2, col, row+2, col+1, delta, fmt["kpi_delta_pos"])
    else:
        ws.merge_range(row+2, col, row+2, col+1, delta, fmt["kpi_delta_neg"])


def _page_setup(ws, orientation="landscape", fit_wide=True):
    """Sett utskriftsoppsett og skjul rutenett for rent canvas."""
    ws.hide_gridlines(2)          # 2 = skjul på skjerm OG ved utskrift
    if orientation == "landscape":
        ws.set_landscape()
    else:
        ws.set_portrait()
    ws.set_paper(9)               # A4
    ws.set_margins(0.5, 0.5, 0.75, 0.75)
    ws.set_header("&L&\"Calibri,Bold\"&10 &F  &C&10 &A &R&10 Side &P / &N")
    ws.set_footer("&L&8 Konfidensielt — kun til intern bruk &R&8 Generert &D")
    if fit_wide:
        ws.fit_to_pages(1, 0)


# ──────────────────────────────────────────────────────────────────────────────
# Ark 1 – Dashbord
# ──────────────────────────────────────────────────────────────────────────────
def _build_dashbord(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(DARK_BLUE)
    ws.set_zoom(90)
    _page_setup(ws)

    all_years   = data["all_years"]
    max_year    = data["max_year"]
    ytd_label   = data["ytd_label"]
    full_years  = data["full_years"]
    fy_vals     = data["fy_vals"]
    ytd_vals    = data["ytd_vals"]
    cagr        = data["cagr"]
    grand_total = data["grand_total"]
    top1_brand  = data["top1_brand"]
    top1_share  = data["top1_share"]
    top3_share  = data["top3_share"]
    ann_sum     = data["annual_summary"]
    hhi         = data.get("hhi", 0)
    abc_brands  = data.get("abc_brands", {})
    gini        = data.get("gini", 0)
    peak_m      = data.get("peak_month", "N/A")
    trough_m    = data.get("trough_month", "N/A")
    seas        = data.get("seasonality", {})
    cr3         = data.get("cr3")
    cr5         = data.get("cr5")

    ws.set_column(0, 0, 3)
    for c in range(1, 15):
        ws.set_column(c, c, 13)
    ws.set_column(15, 18, 6)

    # ── Banner ───────────────────────────────────────────────────────────
    ws.set_row(0, 34); ws.set_row(1, 16); ws.set_row(2, 13)
    ws.merge_range(0, 0, 0, 18,
        f"  {report_name}  |  Nettoomsetningsrapport  v{APP_VERSION}", fmt["title"])
    ws.merge_range(1, 0, 1, 18,
        f"  Ledersammendrag  ·  {min(all_years)}–{max_year}  ·  Alle tall i NOK",
        fmt["subtitle"])
    ws.merge_range(2, 0, 2, 18,
        f"  Generert: {datetime.now().strftime('%d.%m.%Y  %H:%M')}",
        fmt["subtitle"])

    # ── KPI-fliser (7 tiles, 2 kolonner hver = kol 1–14) ─────────────────
    ws.set_row(3, 8); ws.set_row(4, 18)
    ws.merge_range(4, 0, 4, 18, "  NØKKELTALL (KPI)", fmt["section_hdr"])
    ws.set_row(5, 5); ws.set_row(6, 20); ws.set_row(7, 40); ws.set_row(8, 18)

    prev_year = max_year - 1
    ytd_yoy = None
    if ytd_vals.get(prev_year, 0) != 0:
        ytd_yoy = (ytd_vals.get(max_year, 0) - ytd_vals.get(prev_year, 0)) / ytd_vals[prev_year]
    fy_yoy = None
    if fy_vals.get(prev_year, 0) != 0:
        fy_yoy = (fy_vals.get(max_year, 0) - fy_vals[prev_year]) / fy_vals[prev_year]

    kr = 6
    # For the current partial year, fy_yoy compares an incomplete period to a
    # full prior year — always wrong. Show ytd_yoy (same-months comparison) instead.
    partial_year_note = f"YTD {ytd_label}" if full_years else ytd_label
    _write_kpi_tile(ws, fmt, kr,  1, "TOTAL NETTOOMSETNING",              grand_total,                     "kpi_value_nok")
    _write_kpi_tile(ws, fmt, kr,  3, f"FY {max_year} ({partial_year_note})", fy_vals.get(max_year, 0),    "kpi_value_nok", ytd_yoy)
    _write_kpi_tile(ws, fmt, kr,  5, f"Hittil i år {ytd_label}",          ytd_vals.get(max_year, 0),      "kpi_value_nok", ytd_yoy)
    _write_kpi_tile(ws, fmt, kr,  7, "CAGR",                              cagr if cagr is not None else 0, "kpi_value_pct")
    _write_kpi_tile(ws, fmt, kr,  9, "TOPP-VAREMERKE",                    top1_brand,                      "kpi_text")
    _write_kpi_tile(ws, fmt, kr, 11, "TOPP-VAREMERKE ANDEL",              top1_share,                      "kpi_value_pct")
    _write_kpi_tile(ws, fmt, kr, 13, "TOPP 3 ANDEL",                      top3_share,                      "kpi_value_pct")

    # ── Sparkline-strip: 12-mnd trend for hvert år ───────────────────────
    # Trendanalyse-ark: monthly_data_row=5 (0-idx), år i rekkefølge
    spr = kr + 3   # sparkline row
    ws.set_row(spr, 36)
    ws.write(spr, 0, "", fmt["sparkline_cell"])
    # Use actual trend direction based on comparing last two full years
    _trend_arrow = "▲" if (len(full_years) >= 2 and fy_vals.get(full_years[-1], 0) >= fy_vals.get(full_years[-2], 0)) else "▼"
    ws.merge_range(spr, 1, spr, 2, f"{_trend_arrow} 12-mnd trend", fmt["sparkline_label"])

    for yr_i, yr in enumerate(all_years):
        excel_data_row = 5 + yr_i + 1   # 1-indexed: row 6 for first year
        sparkline_range = f'Trendanalyse!B{excel_data_row}:M{excel_data_row}'
        col_for_sparkline = 3 + yr_i * 2   # 2 cols per year
        if col_for_sparkline + 1 <= 14:
            ws.merge_range(spr, col_for_sparkline, spr, col_for_sparkline + 1,
                           "", fmt["sparkline_cell"])
            ws.write(spr, col_for_sparkline,
                     f"{yr}", fmt["sparkline_label"])
            # Sparkline goes in first cell of pair
            try:
                ws.add_sparkline(spr, col_for_sparkline, {
                    "range":          sparkline_range,
                    "type":           "column",
                    "high_point":     True,
                    "low_point":      True,
                    "series_color":   MID_BLUE,
                    "high_color":     ACCENT_GREEN,
                    "low_color":      ACCENT_RED,
                })
            except Exception:
                pass   # ignorerer sparkline-feil (f.eks. for lite data)

    # Fyll resten av sparkline-raden
    for c in range(3 + len(all_years) * 2, 19):
        ws.write(spr, c, "", fmt["sparkline_cell"])

    # ── Ledersammendrag ──────────────────────────────────────────────────
    sum_start = spr + 2
    ws.set_row(sum_start - 1, 8)
    ws.set_row(sum_start, 18)
    ws.merge_range(sum_start, 0, sum_start, 18, "  LEDERSAMMENDRAG", fmt["section_hdr"])

    bullets = []
    bullets.append(
        f"▸  Total nettoomsetning alle år: {_nok(grand_total)}  |  "
        f"{len(all_years)} år med data ({min(all_years)}–{max_year})"
    )
    # Always show YTD comparison for the current year — full-year vs full-year
    # would be misleading because max_year is a partial year (data up to max_month).
    bullets.append(
        f"▸  FY {max_year} ({ytd_label} YTD): {_nok(fy_vals.get(max_year, 0))}"
        + (f"  {'▲' if ytd_yoy >= 0 else '▼'} {abs(ytd_yoy)*100:.1f}% vs {ytd_label} {prev_year}"
           if ytd_yoy is not None else "  (første dataår)")
    )
    if ytd_yoy is not None:
        arrow = "▲" if ytd_yoy >= 0 else "▼"
        bullets.append(
            f"▸  Hittil i år ({ytd_label}) {max_year}: {_nok(ytd_vals.get(max_year, 0))}  "
            f"{arrow} {abs(ytd_yoy)*100:.1f}% vs {ytd_label} {prev_year}  "
            f"(samme periode sammenlignet — ikke full-år mot del-år)"
        )
    if cagr is not None and len(full_years) >= 2:
        cagr_lbl = ("Sterk" if cagr >= 0.15 else
                    "Moderat" if cagr >= 0.05 else
                    "Svak" if cagr >= 0 else "Negativ")
        bullets.append(
            f"▸  CAGR ({full_years[0]}–{full_years[-1]}): {cagr*100:.1f}%  "
            f"— {cagr_lbl} veksttrajektorie over {full_years[-1] - full_years[0]} hele år"
        )
    bullets.append(
        f"▸  Topp-varemerke: {top1_brand}  ({top1_share*100:.1f}% av omsetning)  |  "
        f"Topp 3: {top3_share*100:.1f}% av omsetning"
    )
    if cr3 is not None:
        bullets.append(
            f"▸  Konsentrasjonsrate CR3: {cr3*100:.1f}%"
            + (f"  ·  CR5: {cr5*100:.1f}%" if cr5 else "")
            + "  —  andel topp-3/5 varemerker av total omsetning  (Utton 1975)"
        )
    if not ann_sum.empty and "net_sales" in ann_sum.columns:
        ns = pd.to_numeric(ann_sum["net_sales"], errors="coerce")
        valid = ann_sum[ns.notna()]
        if len(valid):
            best = valid.loc[ns[ns.notna()].idxmax()]
            bullets.append(
                f"▸  Beste år etter omsetning: {best['year']}  ({_nok(float(best['net_sales']))})"
            )
    gini_lbl = ("Svært skjev" if gini > 0.6 else "Moderat skjev" if gini > 0.35 else "Relativt jevn")
    bullets.append(
        f"▸  Gini-koeffisient: {gini:.3f}  —  {gini_lbl} fordeling mellom varemerker  (Gini 1912)"
    )

    ws.set_row(sum_start + 1, 5)
    for i, b in enumerate(bullets):
        ws.set_row(sum_start + 2 + i, 17)
        ws.merge_range(sum_start + 2 + i, 1, sum_start + 2 + i, 18, b, fmt["bullet"])

    last_bul = sum_start + 2 + len(bullets) - 1

    # ── Årssammendrag ────────────────────────────────────────────────────
    ws.set_row(last_bul + 1, 8)
    tbl = last_bul + 2
    ws.set_row(tbl, 18)
    ws.merge_range(tbl, 0, tbl, 8, "  ÅRSSAMMENDRAG NETTOOMSETNING", fmt["section_hdr"])
    tbl += 1
    ws.set_row(tbl, 20)
    ann_hdrs = ["År", "Nettoomsetning (NOK)", "Antall solgt",
                f"ÅoÅ-vekst *",
                f"Hittil i år {ytd_label} (NOK)", f"YTD ÅoÅ ({ytd_label})", "Beste kvartal"]
    ann_widths = [14, 24, 14, 18, 26, 22, 14]
    for ci, (h, w) in enumerate(zip(ann_hdrs, ann_widths)):
        hf = fmt["col_hdr_left"] if ci == 0 else fmt["col_hdr"]
        ws.write(tbl, 1 + ci, h, hf)
        ws.set_column(1 + ci, 1 + ci, w)
    tbl += 1

    for ri, (_, row) in enumerate(ann_sum.iterrows()):
        yr_str = str(row.get("year", ""))
        is_current = yr_str == str(max_year)
        is_alt = ri % 2 == 1
        lf   = fmt["cell_a"]         if is_alt else fmt["cell"]
        nf   = fmt["cell_nok_a"]     if is_alt else fmt["cell_nok"]
        pf   = fmt["cell_pct_yoy_a"] if is_alt else fmt["cell_pct_yoy"]
        cf   = fmt["cell_c_a"]       if is_alt else fmt["cell_c"]
        intf = fmt["cell_int_a"]     if is_alt else fmt["cell_int"]
        ws.set_row(tbl + ri, 17)
        # Label partial year clearly so management can't mistake it for full-year
        label = f"{yr_str} (YTD {ytd_label})" if is_current else yr_str
        ws.write(tbl + ri, 1, label, lf)
        ws.write(tbl + ri, 2, _v(row.get("net_sales", 0)), nf)
        units_v = row.get("units", 0)
        ws.write(tbl + ri, 3, int(units_v) if units_v else 0, intf)
        yoy = _pct(row.get("yoy"))
        # For current year, yoy is None (suppressed in data_processor) — show "—"
        ws.write(tbl + ri, 4, yoy if yoy != "" else "—", pf if yoy != "" else lf)
        ws.write(tbl + ri, 5, _v(row.get("ytd", 0)), nf)
        ytd_y = _pct(row.get("ytd_yoy"))
        ws.write(tbl + ri, 6, ytd_y, pf if ytd_y != "" else lf)
        ws.write(tbl + ri, 7, str(row.get("best_q", "—")), cf)

    # Footnote explaining the asterisk on ÅoÅ-vekst column
    tbl_end = tbl + len(ann_sum)
    ws.set_row(tbl_end, 14)
    ws.merge_range(tbl_end, 1, tbl_end, 7,
        f"  * ÅoÅ-vekst for {max_year} vises ikke — {max_year} er et delvis år ({ytd_label}). "
        f"Bruk kolonnen «YTD ÅoÅ» for en rettferdig sammenligning av samme periode.",
        fmt["note"])

    # ── Strategiske innsikter ────────────────────────────────────────────
    pareto_df = data.get("pareto", pd.DataFrame())
    n_brands_80 = 0
    if not pareto_df.empty and "cumulative" in pareto_df.columns:
        n_brands_80 = int((pareto_df["cumulative"] <= 0.80).sum()) + 1

    n_a = sum(1 for v in abc_brands.values() if v == "A")
    n_b = sum(1 for v in abc_brands.values() if v == "B")
    n_c = sum(1 for v in abc_brands.values() if v == "C")

    hhi_lbl = ("Høy konsentrasjon" if hhi > 0.25 else
               "Moderat konsentrasjon" if hhi > 0.15 else "Diversifisert")
    peak_idx   = seas.get(peak_m, 1.0)
    trough_idx = seas.get(trough_m, 1.0)
    cr_str = ""
    if cr3 is not None:
        cr_str = f"  ·  CR3={cr3*100:.1f}%  CR5={cr5*100:.1f}%" if cr5 else f"  ·  CR3={cr3*100:.1f}%"

    si_row = tbl + len(ann_sum) + 3
    ws.set_row(si_row, 18)
    ws.merge_range(si_row, 0, si_row, 18, "  STRATEGISKE INNSIKTER", fmt["section_hdr"])
    si_row += 1; ws.set_row(si_row, 5); si_row += 1

    insights = [
        ("PORTEFØLJEKONSENTRASJON  (HHI + Konsentrasjonsrate CR3/CR5)",
         f"HHI = {hhi:.3f}  →  {hhi_lbl}{cr_str}.  "
         + ("Høy avhengighet av få varemerker — vurder å bredde porteføljen."
            if hhi > 0.25 else
            "Moderat spredning gir rimelig motstandskraft."
            if hhi > 0.15 else
            "Omsetningen er godt fordelt — lav enkeltmerkeavhengighet.")
         + "  Ref: Herfindahl-Hirschman (1964); Utton (1975)"),
        ("ABC-KLASSIFISERING  (siste hele år)",
         f"A-klasse: {n_a} varemerker (~70% av omsetning)  ·  "
         f"B-klasse: {n_b} (neste 20%)  ·  C-klasse: {n_c} (bunn 10%)  —  "
         "Prioriter ressurser på A-klasse; evaluer C-klasse for rasjonalisering.  "
         "Ref: Dickie (1951)"),
        ("PARETO-ANALYSE  (80/20-regelen)",
         f"{n_brands_80} varemerke(r) → 80% av total omsetning (av {len(abc_brands)} totalt).  "
         + ("Sterk Pareto-konsentrasjon — avhengighetsrisiko er forhøyet."
            if n_brands_80 <= 3 else "Sunn Pareto-fordeling.")
         + "  Ref: Juran & Godfrey (1999)"),
        ("GINI-KOEFFISIENT  (omsetningsulikhet)",
         f"Gini = {gini:.3f}  —  "
         f"{'Svært skjev' if gini > 0.6 else 'Moderat skjev' if gini > 0.35 else 'Relativt jevn'} "
         f"fordeling mellom varemerker.  "
         f"{'Vurder tiltak for å redusere konsentrasjon.' if gini > 0.5 else 'Akseptabelt diversifisert.'}  "
         "Ref: Gini (1912)"),
        ("SESONGMØNSTER",
         f"Toppmåned: {_M_MAP.get(peak_m, peak_m)} (indeks {peak_idx:.2f}×)  ·  "
         f"Bunnmåned: {_M_MAP.get(trough_m, trough_m)} (indeks {trough_idx:.2f}×).  "
         f"Topp/bunn-ratio: {peak_idx / max(trough_idx, 0.01):.1f}×  —  "
         + ("Høy sesongsvingning: planlegg lager og likviditet."
            if peak_idx / max(trough_idx, 0.01) > 2.0 else "Moderat sesongvariasjon.")
         + "  Ref: Makridakis et al. (1998)"),
    ]

    for label, detail in insights:
        ws.set_row(si_row, 16)
        ws.merge_range(si_row, 1, si_row, 18, f"  {label}", fmt["insight_bold"])
        si_row += 1
        ws.set_row(si_row, 30)
        ws.merge_range(si_row, 1, si_row, 18, f"  {detail}", fmt["insight_body"])
        si_row += 1
        ws.set_row(si_row, 5); si_row += 1

    ws.set_row(si_row, 14)
    ws.merge_range(si_row, 1, si_row, 18,
        "  Rammeverk: HHI (1964)  ·  CR3/CR5 (Utton 1975)  ·  ABC (Dickie 1951)  "
        "·  Pareto (Juran & Godfrey 1999)  ·  Gini (1912)  ·  Sesong (Makridakis et al. 1998)",
        fmt["note"])



# ──────────────────────────────────────────────────────────────────────────────
# Ark 2 – Trendanalyse
# ──────────────────────────────────────────────────────────────────────────────
def _build_trendanalyse(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(MID_BLUE)
    ws.set_zoom(90)
    _page_setup(ws)

    all_years     = data["all_years"]
    monthly_pivot = data["monthly_pivot"]
    monthly_yoy   = data["monthly_yoy"]
    qpivot        = data["quarterly_pivot"]
    seas_dict     = data["seasonality"]

    ws.set_column(0, 0, 18)       # År / Periode label
    for c in range(1, 13):
        ws.set_column(c, c, 11)   # month columns
    ws.set_column(13, 13, 14)     # TOTALT
    ws.set_column(14, 16, 11)     # helper / waterfall cols

    ws.set_row(0, 34); ws.set_row(1, 15)
    ws.merge_range(0, 0, 0, 16,
        f"  {report_name}  |  Trendanalyse  —  Nettoomsetning per periode", fmt["title"])
    ws.merge_range(1, 0, 1, 16, "  Alle tall i NOK", fmt["subtitle"])
    ws.set_row(2, 8)

    # ── Månedlig pivot ───────────────────────────────────────────────────
    monthly_hdr_row  = 3
    monthly_col_row  = 4
    monthly_data_row = 5

    ws.set_row(monthly_hdr_row, 18)
    ws.merge_range(monthly_hdr_row, 0, monthly_hdr_row, 13,
                   "  MÅNEDLIG NETTOOMSETNING (NOK)", fmt["section_hdr"])

    ws.set_row(monthly_col_row, 20)
    ws.write(monthly_col_row, 0, "År", fmt["col_hdr_left"])
    for mi, m_en in enumerate(MONTHS_EN):
        ws.write(monthly_col_row, 1 + mi, _M_MAP.get(m_en, m_en), fmt["col_hdr"])
    ws.write(monthly_col_row, 13, "TOTALT", fmt["col_hdr"])

    n_data_years = 0
    r = monthly_data_row
    for ri, year_str in enumerate(monthly_pivot.index):
        is_total = year_str == "TOTAL"
        is_alt   = ri % 2 == 1 and not is_total
        lf = fmt["total"]     if is_total else (fmt["cell_a"]     if is_alt else fmt["cell"])
        nf = fmt["total_nok"] if is_total else (fmt["cell_nok_a"] if is_alt else fmt["cell_nok"])
        ws.set_row(r + ri, 17)
        ws.write(r + ri, 0, year_str, lf)
        for mi, m_en in enumerate(MONTHS_EN):
            v = float(monthly_pivot.loc[year_str, m_en]) if m_en in monthly_pivot.columns else 0
            ws.write(r + ri, 1 + mi, v if v else "",
                     nf if v else (fmt["cell_a"] if is_alt else fmt["cell"]))
        tot = float(monthly_pivot.loc[year_str, "TOTAL"]) if "TOTAL" in monthly_pivot.columns else 0
        ws.write(r + ri, 13, tot, nf)
        if not is_total:
            n_data_years += 1

    ws.freeze_panes(monthly_col_row + 1, 1)

    next_r = r + len(monthly_pivot) + 2

    # Linjediagram — månedlig omsetning per år
    if n_data_years > 0:
        chart = wb.add_chart({"type": "line"})
        colors = [MID_BLUE, ACCENT_GREEN, ACCENT_RED, PURPLE, TEAL, ACCENT_ORG, DARK_BLUE]
        for i in range(n_data_years):
            chart.add_series({
                "name":       str(monthly_pivot.index[i]),
                "categories": ["Trendanalyse", monthly_col_row, 1, monthly_col_row, 12],
                "values":     ["Trendanalyse", monthly_data_row + i, 1, monthly_data_row + i, 12],
                "line":       {"width": 2.5, "color": colors[i % len(colors)]},
                "marker":     {"type": "circle", "size": 5,
                               "fill":   {"color": colors[i % len(colors)]},
                               "border": {"color": WHITE}},
            })
        chart.set_title({"name": "Månedlig nettoomsetning per år (NOK)"})
        chart.set_x_axis({"name": "Måned"})
        chart.set_y_axis({"name": "NOK", "num_format": "#,##0"})
        chart.set_legend({"position": "bottom"})
        chart.set_style(10)
        chart.set_size({"width": 700, "height": 340})
        ws.insert_chart(next_r, 0, chart, {"x_offset": 2, "y_offset": 4})

    next_r += 19

    # ── Månedlig ÅoÅ-vekst ──────────────────────────────────────────────
    ws.set_row(next_r, 18)
    ws.merge_range(next_r, 0, next_r, 13, "  MÅNEDLIG ÅOÅ-VEKST %", fmt["section_hdr"])
    next_r += 1
    ws.set_row(next_r, 20)
    ws.write(next_r, 0, "Periode", fmt["col_hdr_left"])
    for mi, m_en in enumerate(MONTHS_EN):
        ws.write(next_r, 1 + mi, _M_MAP.get(m_en, m_en), fmt["col_hdr"])
    ws.write(next_r, 13, "Hele år", fmt["col_hdr"])
    yoy_hdr_row = next_r
    next_r += 1

    yoy_data_start = next_r
    for ri, (_, row) in enumerate(monthly_yoy.iterrows()):
        is_alt = ri % 2 == 1
        lf = fmt["cell_a"] if is_alt else fmt["cell"]
        pf = fmt["cell_pct_yoy_a"] if is_alt else fmt["cell_pct_yoy"]
        ws.set_row(next_r + ri, 17)
        ws.write(next_r + ri, 0, str(row.get("label", "")), lf)
        for mi, m_en in enumerate(MONTHS_EN):
            v = _pct(row.get(m_en))
            ws.write(next_r + ri, 1 + mi, v, pf if v != "" else lf)
        fy_v = _pct(row.get("Full Year"))
        ws.write(next_r + ri, 13, fy_v, pf if fy_v != "" else lf)

    # Icon set kondisjonell formatering på ÅoÅ-tabellen
    if len(monthly_yoy) > 0:
        ws.conditional_format(
            yoy_data_start, 1,
            yoy_data_start + len(monthly_yoy) - 1, 13,
            {"type": "icon_set", "icon_style": "3_arrows"}
        )

    next_r += max(len(monthly_yoy), 1) + 3

    # ── Vannfalls-diagram: ÅoÅ månedlig endring ──────────────────────────
    # Simuleres som to stolpeserier: pos (grønn) og neg (rød)
    if len(monthly_yoy) > 0 and n_data_years >= 2:
        ws.set_row(next_r, 18)
        ws.merge_range(next_r, 0, next_r, 13,
                       "  ÅOÅ MÅNEDLIG ENDRING — SISTE VS NEST SISTE ÅR (NOK)",
                       fmt["section_hdr"])
        next_r += 1

        # Beregn absolutt månedlig endring — kun for måneder der BEGGE år har data.
        # For det nåværende delåret (max_year) er måneder etter max_month tomme.
        # Å inkludere dem ville gi enorme falske negative søyler i diagrammet.
        y_last    = str(all_years[-1])
        y_prev    = str(all_years[-2])
        max_month = data.get("max_month", 12)
        # Only plot months where the most recent year actually has data
        plot_months = max_month if int(all_years[-1]) == data.get("max_year", 9999) else 12

        delta_pos   = []
        delta_neg   = []
        month_labels = []
        for mi, m_en in enumerate(MONTHS_EN):
            if mi >= plot_months:
                break          # stop at the last month with real data
            v_last = float(monthly_pivot.loc[y_last, m_en]) if (y_last in monthly_pivot.index and m_en in monthly_pivot.columns) else 0
            v_prev = float(monthly_pivot.loc[y_prev, m_en]) if (y_prev in monthly_pivot.index and m_en in monthly_pivot.columns) else 0
            delta = v_last - v_prev
            delta_pos.append(delta if delta > 0 else None)
            delta_neg.append(delta if delta <= 0 else None)
            month_labels.append(MONTHS_NO[mi])

        n_wfall = len(month_labels)

        # Skriv hjelpedata til skjulte kolonner (15, 16, 17)
        wfall_data_row = next_r
        ws.write(wfall_data_row, 14, "Positiv endring (NOK)", fmt["col_hdr"])
        ws.write(wfall_data_row, 15, "Negativ endring (NOK)", fmt["col_hdr"])
        next_r += 1
        for mi in range(n_wfall):
            ws.write(next_r + mi, 0,  month_labels[mi], fmt["cell"])
            ws.write(next_r + mi, 14, delta_pos[mi] if delta_pos[mi] is not None else "", fmt["cell_nok"])
            ws.write(next_r + mi, 15, delta_neg[mi] if delta_neg[mi] is not None else "", fmt["cell_nok"])

        wfall_chart = wb.add_chart({"type": "column"})
        wfall_chart.add_series({
            "name":       f"Vekst {y_last} vs {y_prev}",
            "categories": ["Trendanalyse", next_r, 0, next_r + n_wfall - 1, 0],
            "values":     ["Trendanalyse", next_r, 14, next_r + n_wfall - 1, 14],
            "fill":       {"color": ACCENT_GREEN},
            "border":     {"color": "#145A32"},
        })
        wfall_chart.add_series({
            "name":       f"Nedgang {y_last} vs {y_prev}",
            "categories": ["Trendanalyse", next_r, 0, next_r + n_wfall - 1, 0],
            "values":     ["Trendanalyse", next_r, 15, next_r + n_wfall - 1, 15],
            "fill":       {"color": ACCENT_RED},
            "border":     {"color": "#78281F"},
        })
        ytd_note = f" (YTD {MONTHS_NO[plot_months-1]})" if plot_months < 12 else ""
        wfall_chart.set_title({"name": f"Månedlig omsetningsendring{ytd_note}: {y_last} vs {y_prev} (NOK)"})
        wfall_chart.set_x_axis({"name": "Måned"})
        wfall_chart.set_y_axis({"name": "Endring (NOK)", "num_format": "#,##0"})
        wfall_chart.set_legend({"position": "bottom"})
        wfall_chart.set_style(10)
        wfall_chart.set_size({"width": 700, "height": 320})
        ws.insert_chart(next_r, 0, wfall_chart, {"x_offset": 2, "y_offset": 4})
        next_r += n_wfall + 20  # data rows + chart rows

    # ── Kvartalsoversikt ─────────────────────────────────────────────────
    ws.set_row(next_r, 18)
    ws.merge_range(next_r, 0, next_r, 6, "  KVARTALSVIS NETTOOMSETNING (NOK)", fmt["section_hdr"])
    next_r += 1
    ws.set_row(next_r, 20)
    for ci, h in enumerate(["År", "K1", "K2", "K3", "K4", "Totalt", "K2-andel"]):
        hf = fmt["col_hdr_left"] if ci == 0 else fmt["col_hdr"]
        ws.write(next_r, ci, h, hf)
    next_r += 1

    for ri, (_, row) in enumerate(qpivot.iterrows()):
        is_alt = ri % 2 == 1
        lf = fmt["cell_a"] if is_alt else fmt["cell"]
        nf = fmt["cell_nok_a"] if is_alt else fmt["cell_nok"]
        pf = fmt["cell_pct_a"] if is_alt else fmt["cell_pct"]
        ws.set_row(next_r + ri, 17)
        ws.write(next_r + ri, 0, str(row.get("År", "")), lf)
        for qi, q in enumerate(["K1", "K2", "K3", "K4"]):
            v = row.get(q, 0)
            ws.write(next_r + ri, 1 + qi, v if v else "", nf if v else lf)
        tot = row.get("Total", 0)
        ws.write(next_r + ri, 5, tot if tot else "", nf if tot else lf)
        q2s = _pct(row.get("K2-andel"))
        ws.write(next_r + ri, 6, q2s, pf if q2s != "" else lf)

    next_r += len(qpivot) + 3

    # ── Sesongindeks ─────────────────────────────────────────────────────
    ws.set_row(next_r, 18)
    ws.merge_range(next_r, 0, next_r, 13,
                   "  SESONGINDEKS  (1,00 = gjennomsnittsmåned, kun hele kalenderår)",
                   fmt["section_hdr"])
    next_r += 1
    ws.set_row(next_r, 20)
    ws.write(next_r, 0, "Måned", fmt["col_hdr_left"])
    for mi, m_en in enumerate(MONTHS_EN):
        ws.write(next_r, 1 + mi, _M_MAP.get(m_en, m_en), fmt["col_hdr"])
    next_r += 1
    ws.set_row(next_r, 20)
    ws.write(next_r, 0, "Indeks", fmt["cell"])
    for mi, m_en in enumerate(MONTHS_EN):
        v = seas_dict.get(m_en, 1.0)
        sf = fmt["seas_hi"] if v >= 1.1 else (fmt["seas_lo"] if v <= 0.9 else fmt["seas_mid"])
        ws.write(next_r, 1 + mi, v, sf)
    next_r += 2
    ws.merge_range(next_r, 0, next_r, 13,
        "  Grønt ≥ 1,10 (sesongtopp)  ·  Rødt ≤ 0,90 (sesongbunn)  ·  "
        "Ref: Makridakis, Wheelwright & Hyndman (1998)",
        fmt["note"])


# ──────────────────────────────────────────────────────────────────────────────
# Ark 3 – Varemerkeanalyse
# ──────────────────────────────────────────────────────────────────────────────
def _build_varemerkeanalyse(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(ACCENT_GREEN)
    ws.set_zoom(90)
    _page_setup(ws)

    all_years  = data["all_years"]
    brand_perf = data["brand_perf"]
    pareto_df  = data["pareto"]
    last_full  = data["last_full_year"]
    max_year   = data["max_year"]
    ytd_label  = data["ytd_label"]
    abc_brands = data.get("abc_brands", {})

    n_fy  = len(all_years)
    n_yoy = max(0, len(all_years) - 1)
    n_cols = 1 + n_fy + 1 + n_yoy + 1 + 1

    ws.set_column(0, 0, 28)                          # brand name
    for c in range(1, 1 + n_fy):
        ws.set_column(c, c, 16)                      # FY columns
    ws.set_column(1 + n_fy, 1 + n_fy, 16)           # YTD column
    for c in range(2 + n_fy, 2 + n_fy + n_yoy - 1):
        ws.set_column(c, c, 14)                      # prior-year YoY columns
    if n_yoy > 0:
        ws.set_column(1 + n_fy + n_yoy, 1 + n_fy + n_yoy, 22)  # current YTD YoY (wider for label)
    ws.set_column(2 + n_fy + n_yoy, 2 + n_fy + n_yoy, 14)      # share column
    ws.set_column(3 + n_fy + n_yoy, 3 + n_fy + n_yoy, 8)       # ABC

    ws.set_row(0, 34); ws.set_row(1, 15)
    ws.merge_range(0, 0, 0, n_cols,
        f"  {report_name}  |  Varemerkeanalyse  —  Nettoomsetning per varemerke",
        fmt["title"])
    ws.merge_range(1, 0, 1, n_cols,
        "  Alle tall i NOK  ·  Sortert etter siste hele år  ·  "
        "ABC = porteføljeklassifisering (Dickie 1951)  ·  "
        "Datalinjer = FY-omsetning  ·  Piler = ÅoÅ-retning",
        fmt["subtitle"])
    ws.set_row(2, 8)
    ws.set_row(3, 18)
    ws.merge_range(3, 0, 3, n_cols, "  VAREMERKEANALYSE", fmt["section_hdr"])

    ws.set_row(4, 22)

    def _yoy_header(y1, y0, max_yr, ytd_lbl):
        if y1 == max_yr:
            return f"YTD ÅoÅ {y1}v{y0} ({ytd_lbl})"
        return f"ÅoÅ {y1}v{y0}"

    headers = (["Varemerke"] + [f"FY {y}" for y in all_years] +
               [f"Hittil i år {ytd_label}"] +
               [_yoy_header(all_years[i], all_years[i-1], max_year, ytd_label)
                for i in range(1, len(all_years))] +
               [f"Andel FY{last_full}", "ABC"])
    for ci, h in enumerate(headers):
        hf = fmt["col_hdr_left"] if ci == 0 else fmt["col_hdr"]
        ws.write(4, ci, h, hf)

    ws.freeze_panes(5, 1)

    n_brands_data = len([r for _, r in brand_perf.iterrows()
                          if str(r.get("brand", "")).upper() != "TOTAL"])
    data_start_row = 5

    for ri, (_, row) in enumerate(brand_perf.iterrows()):
        brand_name = str(row.get("brand", ""))
        is_total   = brand_name.upper() == "TOTAL"
        is_alt     = ri % 2 == 1 and not is_total
        lf  = fmt["total"]         if is_total else (fmt["cell_a"]         if is_alt else fmt["cell"])
        nf  = fmt["total_nok"]     if is_total else (fmt["cell_nok_a"]     if is_alt else fmt["cell_nok"])
        pf  = fmt["total_pct_yoy"] if is_total else (fmt["cell_pct_yoy_a"] if is_alt else fmt["cell_pct_yoy"])
        sf  = fmt["total_pct"]     if is_total else (fmt["cell_pct_a"]     if is_alt else fmt["cell_pct"])
        ws.set_row(5 + ri, 17)

        ci = 0
        ws.write(5 + ri, ci, brand_name, lf); ci += 1
        for y in all_years:
            v = _v(row.get(f"FY{y}", 0))
            ws.write(5 + ri, ci, v, nf if v != "" and v else lf); ci += 1
        ytd_v = _v(row.get("YTD", 0))
        ws.write(5 + ri, ci, ytd_v, nf if ytd_v != "" and ytd_v else lf); ci += 1
        for i in range(1, len(all_years)):
            y0, y1 = all_years[i-1], all_years[i]
            yoy_v = _pct(row.get(f"YoY_{y1}v{y0}"))
            ws.write(5 + ri, ci, yoy_v if yoy_v != "" else "—",
                     pf if yoy_v != "" else lf); ci += 1
        share_v = _pct(row.get("share"))
        ws.write(5 + ri, ci, share_v, sf if share_v != "" else lf); ci += 1
        abc_class = abc_brands.get(brand_name, "")
        if is_total:
            ws.write(5 + ri, ci, "—", fmt["total_c"])
        elif abc_class in ("A", "B", "C"):
            ws.write(5 + ri, ci, abc_class, fmt[f"abc_{abc_class}"])
        else:
            ws.write(5 + ri, ci, "", lf)

    # ── Kondisjonell formatering ─────────────────────────────────────────
    # Datalinjer på siste hele år-kolonne (kolonne index = 1 + n_fy-1 = n_fy)
    if n_brands_data > 0:
        fy_col = n_fy   # 0-indexed column for last full year
        ws.conditional_format(
            data_start_row, fy_col,
            data_start_row + n_brands_data - 1, fy_col,
            {
                "type":             "data_bar",
                "bar_color":        MID_BLUE,
                "bar_border_color": DARK_BLUE,
                "data_bar_2010":    True,
            }
        )

        # Pil-ikonset på ÅoÅ-kolonner
        for i in range(n_yoy):
            yoy_col = 2 + n_fy + i
            ws.conditional_format(
                data_start_row, yoy_col,
                data_start_row + n_brands_data - 1, yoy_col,
                {"type": "icon_set", "icon_style": "3_arrows"}
            )

    pareto_section = 5 + len(brand_perf) + 3

    # ── Pareto-analyse ───────────────────────────────────────────────────
    ws.set_row(pareto_section, 18)
    ws.merge_range(pareto_section, 0, pareto_section, 5,
                   "  PARETO-ANALYSE  —  80/20 Omsetningskonsentrasjon  (Juran & Godfrey 1999)",
                   fmt["section_hdr"])
    pareto_section += 1
    ws.set_row(pareto_section, 20)
    pareto_hdrs = ["Rang", "Varemerke", "Total omsetning (NOK)", "% av totalt", "Kumulativ %", ""]
    for ci, h in enumerate(pareto_hdrs):
        hf = fmt["col_hdr_left"] if ci == 1 else fmt["col_hdr"]
        ws.write(pareto_section, ci, h, hf)
    pareto_data_start = pareto_section + 1
    pareto_section += 1

    for ri, (_, row) in enumerate(pareto_df.iterrows()):
        is_80  = str(row.get("threshold_80", "")) == "80%"
        is_alt = ri % 2 == 1 and not is_80
        lf  = fmt["p80_c"]   if is_80 else (fmt["cell_c_a"]   if is_alt else fmt["cell_c"])
        bf  = fmt["p80_l"]   if is_80 else (fmt["cell_a"]     if is_alt else fmt["cell"])
        nf  = fmt["p80_nok"] if is_80 else (fmt["cell_nok_a"] if is_alt else fmt["cell_nok"])
        pf  = fmt["p80_pct"] if is_80 else (fmt["cell_pct_a"] if is_alt else fmt["cell_pct"])
        tf  = fmt["p80_tag"] if is_80 else lf
        ws.set_row(pareto_section + ri, 17)
        ws.write(pareto_section + ri, 0, int(row.get("rank", ri+1)),         lf)
        ws.write(pareto_section + ri, 1, str(row.get("brand", "")),          bf)
        ws.write(pareto_section + ri, 2, float(row.get("total_revenue", 0)), nf)
        ws.write(pareto_section + ri, 3, float(row.get("pct", 0)),           pf)
        ws.write(pareto_section + ri, 4, float(row.get("cumulative", 0)),    pf)
        ws.write(pareto_section + ri, 5, str(row.get("threshold_80", "")),   tf)

    # Kumulativ datalinjeformatering på Pareto-tabellen
    if len(pareto_df) > 0:
        ws.conditional_format(
            pareto_data_start, 2,
            pareto_data_start + len(pareto_df) - 1, 2,
            {"type": "data_bar", "bar_color": ACCENT_GREEN, "data_bar_2010": True}
        )

    n_chart = min(10, len(pareto_df))
    if n_chart > 0:
        bar_chart = wb.add_chart({"type": "bar"})
        bar_chart.add_series({
            "name":       "Nettoomsetning (NOK)",
            "categories": ["Varemerkeanalyse", pareto_data_start, 1,
                           pareto_data_start + n_chart - 1, 1],
            "values":     ["Varemerkeanalyse", pareto_data_start, 2,
                           pareto_data_start + n_chart - 1, 2],
            "fill":       {"color": MID_BLUE},
            "border":     {"color": DARK_BLUE},
        })
        bar_chart.set_title({"name": f"Topp {n_chart} varemerker etter nettoomsetning"})
        bar_chart.set_x_axis({"name": "NOK", "num_format": "#,##0"})
        bar_chart.set_y_axis({"name": "Varemerke", "reverse": True})
        bar_chart.set_legend({"none": True})
        bar_chart.set_style(10)
        bar_chart.set_size({"width": 580, "height": 400})
        chart_row = pareto_section + len(pareto_df) + 2
        ws.insert_chart(chart_row, 0, bar_chart, {"x_offset": 2, "y_offset": 4})


# ──────────────────────────────────────────────────────────────────────────────
# Ark 4 – XYZ-analyse
# ──────────────────────────────────────────────────────────────────────────────
def _build_xyz_analyse(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(TEAL)
    ws.set_zoom(90)
    _page_setup(ws)

    xyz_df         = data.get("xyz_df")
    abc_xyz_matrix = data.get("abc_xyz_matrix", {})

    if xyz_df is None or (hasattr(xyz_df, "empty") and xyz_df.empty):
        ws.merge_range(0, 0, 0, 7,
                       f"  {report_name}  |  XYZ-analyse — Ikke nok data", fmt["title"])
        return

    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 18)
    ws.set_column(2, 2, 18)
    ws.set_column(3, 3, 14)
    ws.set_column(4, 4, 14)
    ws.set_column(5, 5, 9)
    ws.set_column(6, 6, 9)
    ws.set_column(7, 7, 12)

    ws.set_row(0, 34); ws.set_row(1, 15)
    ws.merge_range(0, 0, 0, 7,
        f"  {report_name}  |  XYZ-analyse  —  Etterspørselsvariabilitet per varemerke",
        fmt["title"])
    ws.merge_range(1, 0, 1, 7,
        "  CV = std/gjennomsnitt av månedlig omsetning  ·  "
        "Ref: Silver, Pyke & Peterson (1998); Scholz-Reiter et al. (2012)",
        fmt["subtitle"])
    ws.set_row(2, 8)
    ws.set_row(3, 18)
    ws.merge_range(3, 0, 3, 7, "  METODIKK OG KLASSIFISERING", fmt["section_hdr"])

    exp_rows = [
        "  X = Stabil etterspørsel            (CV < 0,50)  —  Lav variabilitet, forutsigbar omsetning",
        "  Y = Variabel etterspørsel           (0,50 ≤ CV < 1,00)  —  Moderat variabilitet, sesong- el. trendpåvirket",
        "  Z = Svært variabel etterspørsel     (CV ≥ 1,00)  —  Høy variabilitet, vanskelig å forutsi",
    ]
    xyz_colors = [TEAL, "#B7950B", ACCENT_RED]
    for i, (txt, col) in enumerate(zip(exp_rows, xyz_colors)):
        exp_fmt = wb.add_format({
            "font_size": 10, "font_name": "Calibri", "font_color": col, "bold": True,
            "align": "left", "valign": "vcenter", "indent": 1,
        })
        ws.set_row(4 + i, 17)
        ws.merge_range(4 + i, 0, 4 + i, 7, txt, exp_fmt)
    ws.set_row(7, 8)
    ws.set_row(8, 18)
    ws.merge_range(8, 0, 8, 7, "  XYZ-KLASSIFISERING PER VAREMERKE", fmt["section_hdr"])
    ws.set_row(9, 20)
    hdrs = ["Varemerke", "Gj.snitt månedlig (NOK)", "Std.avvik (NOK)",
            "Var.koeff. (CV)", "Antall måneder", "XYZ", "ABC", "Kombinert"]
    for ci, h in enumerate(hdrs):
        hf = fmt["col_hdr_left"] if ci == 0 else fmt["col_hdr"]
        ws.write(9, ci, h, hf)

    ws.freeze_panes(10, 1)
    xyz_fmt_map = {"X": "xyz_X", "Y": "xyz_Y", "Z": "xyz_Z"}

    for ri, (_, row) in enumerate(xyz_df.iterrows()):
        is_alt = ri % 2 == 1
        lf   = fmt["cell_a"]     if is_alt else fmt["cell"]
        nf   = fmt["cell_nok_a"] if is_alt else fmt["cell_nok"]
        cf   = fmt["cell_2dec_a"] if is_alt else fmt["cell_2dec"]
        intf = fmt["cell_int_a"] if is_alt else fmt["cell_int"]
        ws.set_row(10 + ri, 17)
        ws.write(10 + ri, 0, str(row.get("brand", "")),         lf)
        ws.write(10 + ri, 1, float(row.get("mean_monthly", 0)), nf)
        ws.write(10 + ri, 2, float(row.get("std_monthly", 0)),  nf)
        ws.write(10 + ri, 3, float(row.get("cv", 0)),           cf)
        ws.write(10 + ri, 4, int(row.get("n_months", 0)),       intf)
        xyz_k = str(row.get("xyz", ""))
        ws.write(10 + ri, 5, xyz_k,
                 fmt[xyz_fmt_map[xyz_k]] if xyz_k in xyz_fmt_map else lf)
        abc_k = str(row.get("abc", "—"))
        ws.write(10 + ri, 6, abc_k,
                 fmt[f"abc_{abc_k}"] if abc_k in ("A", "B", "C") else lf)
        ws.write(10 + ri, 7, str(row.get("combined", "—")), lf)

    # Datalinjer på CV-kolonnen
    if len(xyz_df) > 0:
        ws.conditional_format(10, 3, 10 + len(xyz_df) - 1, 3,
                               {"type": "data_bar", "bar_color": ACCENT_ORG, "data_bar_2010": True})

    matrix_row = 10 + len(xyz_df) + 3
    ws.set_row(matrix_row, 18)
    ws.merge_range(matrix_row, 0, matrix_row, 4,
                   "  ABC–XYZ KOMBINASJONSMATRISE  (antall varemerker per kvadrant)",
                   fmt["section_hdr"])
    matrix_row += 1

    hdr_fmt = wb.add_format({"bold": True, "font_size": 11, "font_name": "Calibri",
                              "align": "center", "valign": "vcenter", "border": 2,
                              "border_color": DARK_BLUE, "bg_color": DARK_BLUE, "font_color": WHITE})
    ws.set_row(matrix_row, 22)
    ws.write(matrix_row, 0, "", hdr_fmt)
    ws.write(matrix_row, 1, "X  (stabil)",   fmt["xyz_X"])
    ws.write(matrix_row, 2, "Y  (variabel)", fmt["xyz_Y"])
    ws.write(matrix_row, 3, "Z  (erratisk)", fmt["xyz_Z"])
    matrix_row += 1

    for abc_k in ["A", "B", "C"]:
        ws.set_row(matrix_row, 24)
        ws.write(matrix_row, 0, f"{abc_k}-klasse", fmt[f"abc_{abc_k}"])
        for xi, xyz_k in enumerate(["X", "Y", "Z"]):
            cnt = abc_xyz_matrix.get(abc_k, {}).get(xyz_k, 0)
            cell_fmt = wb.add_format({
                "bold": True, "font_size": 14, "font_name": "Calibri",
                "align": "center", "valign": "vcenter",
                "border": 1, "border_color": MID_GREY,
                "bg_color": LIGHT_BLUE if cnt > 0 else LIGHT_GREY,
                "font_color": DARK_BLUE,
            })
            ws.write(matrix_row, 1 + xi, cnt, cell_fmt)
        matrix_row += 1

    matrix_row += 1
    ws.set_row(matrix_row, 32)
    ws.merge_range(matrix_row, 0, matrix_row, 7,
        "  AX = Høy verdi, stabil (prioriter)  ·  AZ = Høy verdi, erratisk (kritisk risiko, øk sikkerhetslager)  ·  "
        "CX = Lav verdi, stabil (standard)  ·  CZ = Lav verdi, erratisk (vurder eliminering)  "
        "·  Ref: Silver, Pyke & Peterson (1998); Scholz-Reiter et al. (2012)",
        fmt["note"])


# ──────────────────────────────────────────────────────────────────────────────
# Ark 5 – Portefølje  (BCG scatter-diagram med fargede serier per kategori)
# ──────────────────────────────────────────────────────────────────────────────
def _build_portfolje(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(PURPLE)
    ws.set_zoom(90)
    _page_setup(ws)

    portfolio_df  = data.get("portfolio_df")
    ref_years     = data.get("portfolio_ref_years")
    avg_growth    = data.get("portfolio_avg_growth")
    avg_share     = data.get("portfolio_avg_share")

    has_data = (portfolio_df is not None and
                hasattr(portfolio_df, "empty") and
                not portfolio_df.empty)

    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 16)
    ws.set_column(2, 2, 16)
    ws.set_column(3, 3, 20)
    ws.set_column(4, 4, 9)
    ws.set_column(5, 5, 18)
    ws.set_column(6, 8, 14)

    ws.set_row(0, 34); ws.set_row(1, 15)
    ref_str = f"{ref_years[0]}–{ref_years[1]}" if ref_years else "N/A"
    ws.merge_range(0, 0, 0, 8,
        f"  {report_name}  |  Porteføljeanalyse  —  BCG-inspirert vekst/andel-matrise",
        fmt["title"])
    ws.merge_range(1, 0, 1, 8,
        f"  Referanseperiode: {ref_str}  ·  Vekstrate = ÅoÅ nettoomsetning  ·  "
        "Andel = andel av total porteføljeomsetning  ·  Ref: Henderson (1970)",
        fmt["subtitle"])
    ws.set_row(2, 8)
    ws.set_row(3, 18)
    ws.merge_range(3, 0, 3, 8, "  KATEGORIER OG STRATEGISKE ANBEFALINGER", fmt["section_hdr"])

    cat_rows = [
        ("Stjerne  ★",       GOLD_BG,      DARK_BLUE,
         "Høy vekst + høy andel  —  Invester for å opprettholde posisjon og utnytte momentum"),
        ("Melkeku  ◆",       ACCENT_GREEN, WHITE,
         "Lav vekst + høy andel  —  Optimaliser margin og bruk overskudd til å finansiere vekst"),
        ("Spørsmålstegn  ?", MID_BLUE,     WHITE,
         "Høy vekst + lav andel  —  Vurder selektivt: øk investering på lovende merkevarer"),
        ("Hund  ✕",          DARK_GREY,    WHITE,
         "Lav vekst + lav andel  —  Evaluer for rasjonalisering eller avvikling"),
    ]
    for i, (cat_name, bg, fc, desc) in enumerate(cat_rows):
        cat_hdr_fmt = wb.add_format({"bold": True, "font_size": 10, "font_name": "Calibri",
                                      "font_color": fc, "bg_color": bg, "align": "left",
                                      "valign": "vcenter", "indent": 1, "border": 1,
                                      "border_color": MID_GREY})
        cat_desc_fmt = wb.add_format({"font_size": 10, "font_name": "Calibri",
                                       "align": "left", "valign": "vcenter", "indent": 1,
                                       "border": 1, "border_color": MID_GREY})
        ws.set_row(4 + i, 17)
        ws.write(4 + i, 0, f"  {cat_name}", cat_hdr_fmt)
        ws.merge_range(4 + i, 1, 4 + i, 8, desc, cat_desc_fmt)
    ws.set_row(8, 8)

    if not has_data:
        ws.set_row(9, 18)
        ws.merge_range(9, 0, 9, 8,
                       "  Ikke nok historiske år (krever minst 2 hele kalenderår)", fmt["note"])
        return

    # ── Porteføljeoversikt ───────────────────────────────────────────────
    ws.set_row(9, 18)
    ws.merge_range(9, 0, 9, 8, f"  PORTEFØLJEOVERSIKT  —  {ref_str}", fmt["section_hdr"])
    ws.set_row(10, 20)
    hdrs = ["Varemerke", f"Omsetningsandel FY{ref_years[1] if ref_years else ''}",
            "Vekstrate ÅoÅ", "Kategori", "ABC"]
    for ci, h in enumerate(hdrs):
        hf = fmt["col_hdr_left"] if ci == 0 else fmt["col_hdr"]
        ws.write(10, ci, h, hf)

    ws.freeze_panes(11, 1)

    cat_fmt_map = {
        "Stjerne":       "cat_stjerne",
        "Melkeku":       "cat_melkeku",
        "Spørsmålstegn": "cat_spm",
        "Hund":          "cat_hund",
        "Ny":            "cat_ny",
    }

    for ri, (_, row) in enumerate(portfolio_df.iterrows()):
        is_alt = ri % 2 == 1
        lf  = fmt["cell_a"] if is_alt else fmt["cell"]
        pf  = fmt["cell_pct_a"] if is_alt else fmt["cell_pct"]
        pyf = fmt["cell_pct_yoy_a"] if is_alt else fmt["cell_pct_yoy"]
        ws.set_row(11 + ri, 17)
        ws.write(11 + ri, 0, str(row.get("brand", "")), lf)
        ws.write(11 + ri, 1, _v(row.get("share_pct", 0)), pf)
        g = _pct(row.get("growth_pct"))
        ws.write(11 + ri, 2, g, pyf if g != "" else lf)
        cat = str(row.get("category", ""))
        ws.write(11 + ri, 3, cat,
                 fmt[cat_fmt_map[cat]] if cat in cat_fmt_map else lf)
        abc_k = str(row.get("abc", "—"))
        ws.write(11 + ri, 4, abc_k,
                 fmt[f"abc_{abc_k}"] if abc_k in ("A", "B", "C") else lf)

    note_row = 11 + len(portfolio_df) + 2
    ws.set_row(note_row, 28)
    ws.merge_range(note_row, 0, note_row, 8,
        f"  Terskelverdi vekst: {avg_growth*100:.1f}%  ·  "
        f"Terskelverdi andel: {avg_share*100:.2f}%  ·  "
        "Interne terskler — relative til denne portefølje, ikke markedet  ·  "
        "Ref: Henderson (1970) — BCG Growth-Share Matrix",
        fmt["note"])

    # ── BCG Scatter-diagram med fargede serier per kategori ──────────────
    # Skriv hjelpedata til høyre (kol 10+) og lag scatter per kategori
    helper_col_base = 10
    ws.set_column(helper_col_base, helper_col_base + 10, 0)   # skjul hjelpekolonner

    cats_present = portfolio_df["category"].unique().tolist()
    chart_data_rows = {}   # category → (start_row, n_rows)

    helper_row = 0
    ws.write(helper_row, helper_col_base,     "Kategori",       fmt["col_hdr"])
    ws.write(helper_row, helper_col_base + 1, "Andel",          fmt["col_hdr"])
    ws.write(helper_row, helper_col_base + 2, "Vekstrate",      fmt["col_hdr"])
    ws.write(helper_row, helper_col_base + 3, "Varemerkenavn",  fmt["col_hdr"])
    helper_row += 1

    for cat in cats_present:
        cat_df = portfolio_df[portfolio_df["category"] == cat]
        start = helper_row
        for _, row in cat_df.iterrows():
            share  = row.get("share_pct")
            growth = row.get("growth_pct")
            if share is None or growth is None:
                continue
            try:
                share_v  = float(share)
                growth_v = float(growth)
            except (TypeError, ValueError):
                continue
            ws.write(helper_row, helper_col_base,     cat,                      fmt["cell"])
            ws.write(helper_row, helper_col_base + 1, share_v,                  fmt["cell_2dec"])
            ws.write(helper_row, helper_col_base + 2, growth_v,                 fmt["cell_2dec"])
            ws.write(helper_row, helper_col_base + 3, str(row.get("brand", "")), fmt["cell"])
            helper_row += 1
        chart_data_rows[cat] = (start, helper_row - start)   # actual written rows

    # Only create the chart object when we know at least one series has data.
    # xlsxwriter raises EmptyChartSeries on wb.close() for any registered chart
    # with no series, even if the chart was never inserted into a sheet.
    has_bcg_data = any(n_r > 0 for _, n_r in chart_data_rows.values())

    chart_insert_row = note_row + 3
    if has_bcg_data:
        bcg_chart = wb.add_chart({"type": "scatter", "subtype": "straight_with_markers"})
        for cat in cats_present:
            if cat not in chart_data_rows:
                continue
            start_r, n_r = chart_data_rows[cat]
            if n_r == 0:
                continue
            color = _BCG_COLORS.get(cat, MID_BLUE)
            marker_type = _BCG_MARKERS.get(cat, "circle")
            bcg_chart.add_series({
                "name":       cat,
                # xlsxwriter scatter: "categories" = x-axis, "values" = y-axis
                "categories": ["Portefølje", start_r, helper_col_base + 1,
                                start_r + n_r - 1, helper_col_base + 1],
                "values":     ["Portefølje", start_r, helper_col_base + 2,
                                start_r + n_r - 1, helper_col_base + 2],
                "marker": {
                    "type":   marker_type,
                    "size":   9,
                    "fill":   {"color": color},
                    "border": {"color": DARK_BLUE},
                },
                "line": {"none": True},
            })
        bcg_chart.set_title({"name": f"BCG-inspirert Vekst/Andel-matrise  ({ref_str})"})
        bcg_chart.set_x_axis({
            "name":       "Omsetningsandel (%) →  Høy andel = høyre",
            "num_format": "0.0%",
            "min": 0,
        })
        bcg_chart.set_y_axis({
            "name":       "Vekstrate (%) ↑  Høy vekst = topp",
            "num_format": "0.0%",
        })
        bcg_chart.set_legend({"position": "bottom"})
        bcg_chart.set_style(10)
        bcg_chart.set_size({"width": 600, "height": 440})
        ws.insert_chart(chart_insert_row, 0, bcg_chart, {"x_offset": 2, "y_offset": 4})

    # Matriseoversikt
    matrix_row = chart_insert_row + 28
    ws.set_row(matrix_row, 18)
    ws.merge_range(matrix_row, 0, matrix_row, 4,
                   "  MATRISEOVERSIKT  (antall varemerker per kvadrant)",
                   fmt["section_hdr"])
    matrix_row += 1

    cats_all = ["Stjerne", "Melkeku", "Spørsmålstegn", "Hund", "Ny"]
    counts = {c: 0 for c in cats_all}
    for _, row in portfolio_df.iterrows():
        cat = str(row.get("category", ""))
        if cat in counts:
            counts[cat] += 1

    ws.set_row(matrix_row, 22)
    for ci, cat in enumerate(cats_all):
        ws.write(matrix_row, ci, cat,
                 fmt[cat_fmt_map[cat]] if cat in cat_fmt_map else fmt["cell_c"])
    matrix_row += 1
    ws.set_row(matrix_row, 24)
    cnt_fmt = wb.add_format({"bold": True, "font_size": 14, "font_name": "Calibri",
                              "align": "center", "valign": "vcenter",
                              "border": 1, "border_color": MID_GREY, "bg_color": LIGHT_BLUE})
    for ci, cat in enumerate(cats_all):
        ws.write(matrix_row, ci, counts[cat], cnt_fmt)


# ──────────────────────────────────────────────────────────────────────────────
# Ark 6 – Topp-artikler  (med år-for-år-kolonner)
# ──────────────────────────────────────────────────────────────────────────────
def _build_topp_artikler(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(GOLD_BORDER)
    ws.set_zoom(85)
    _page_setup(ws)

    top5_brands         = data.get("top5_brands", [])
    top_items_per_brand = data.get("top_items_per_brand", {})
    abc_brands          = data.get("abc_brands", {})
    all_years           = data.get("all_years", [])

    yr_cols_start    = 3
    n_yr_col_pairs   = len(all_years)
    total_antall_col = yr_cols_start + n_yr_col_pairs * 2
    total_nok_col    = total_antall_col + 1
    last_col         = total_nok_col

    ws.set_column(0, 0, 6)
    ws.set_column(1, 1, 16)
    ws.set_column(2, 2, 44)
    for i in range(n_yr_col_pairs):
        ws.set_column(yr_cols_start + i*2,     yr_cols_start + i*2,     12)
        ws.set_column(yr_cols_start + i*2 + 1, yr_cols_start + i*2 + 1, 15)
    ws.set_column(total_antall_col, total_antall_col, 13)
    ws.set_column(total_nok_col,    total_nok_col,    16)

    ws.set_row(0, 34); ws.set_row(1, 15)
    ws.merge_range(0, 0, 0, last_col,
        f"  {report_name}  |  Bestselgende artikler etter antall  —  Topp 5 varemerker",
        fmt["title"])
    ws.merge_range(1, 0, 1, last_col,
        "  Rangert etter totalt antall solgte enheter  ·  "
        "Farge/størrelsesvarianter konsolidert per basisartikkel  ·  "
        "År-for-år-fordeling for trendanalyse",
        fmt["subtitle"])

    r = 2
    for brand in top5_brands:
        items = top_items_per_brand.get(brand)
        if items is None or len(items) == 0:
            continue

        abc_class  = abc_brands.get(brand, "")
        abc_suffix = f"  [{abc_class}-klasse]" if abc_class else ""

        ws.set_row(r, 6); r += 1
        ws.set_row(r, 22)
        ws.merge_range(r, 0, r, last_col, f"  {brand}{abc_suffix}", fmt["section_hdr"])
        r += 1

        ws.set_row(r, 32)
        ws.write(r, 0, "Rang",       fmt["col_hdr"])
        ws.write(r, 1, "Art.nr.",    fmt["col_hdr"])
        ws.write(r, 2, "Beskrivelse", fmt["col_hdr_left"])
        for i, yr in enumerate(all_years):
            c_ant = yr_cols_start + i * 2
            ws.write(r, c_ant,     f"{yr}\nAntall", fmt["col_hdr_yr"])
            ws.write(r, c_ant + 1, f"{yr}\nNOK",    fmt["col_hdr_yr"])
        ws.write(r, total_antall_col, "Total\nAntall", fmt["col_hdr"])
        ws.write(r, total_nok_col,    "Total\nNOK",    fmt["col_hdr"])
        r += 1

        data_start_r = r
        for ri, (_, item) in enumerate(items.iterrows()):
            is_alt = ri % 2 == 1
            lf   = fmt["cell_a"]     if is_alt else fmt["cell"]
            nf   = fmt["cell_nok_a"] if is_alt else fmt["cell_nok"]
            intf = fmt["cell_int_a"] if is_alt else fmt["cell_int"]
            rf   = fmt["rank_gold"]  if ri == 0 else fmt["rank_std"]

            ws.set_row(r + ri, 17)
            ws.write(r + ri, 0, ri + 1,                           rf)
            ws.write(r + ri, 1, str(item.get("article_no", "")),  lf)
            ws.write(r + ri, 2, str(item.get("article_desc", "")), lf)
            for i, yr in enumerate(all_years):
                c_ant = yr_cols_start + i * 2
                u_val = item.get(f"units_{yr}", 0)
                s_val = item.get(f"sales_{yr}", 0.0)
                ws.write(r + ri, c_ant,     int(u_val) if u_val else "",   intf if u_val else fmt["cell_empty_yr"])
                ws.write(r + ri, c_ant + 1, float(s_val) if s_val else "", nf   if s_val else fmt["cell_empty_yr"])
            ws.write(r + ri, total_antall_col, int(item.get("total_units", 0)),   intf)
            ws.write(r + ri, total_nok_col,    float(item.get("total_sales", 0)), nf)

        # Datalinjer på Total NOK-kolonnen for dette varemerket
        if len(items) > 0:
            ws.conditional_format(
                data_start_r, total_nok_col,
                data_start_r + len(items) - 1, total_nok_col,
                {"type": "data_bar", "bar_color": GOLD_BORDER, "data_bar_2010": True}
            )

        r += len(items)

    r += 2
    ws.set_row(r, 16)
    ws.merge_range(r, 0, r, last_col,
        "  Varemerkerangering etter all-time nettoomsetning  ·  "
        "ABC basert på siste hele år (Dickie 1951)  ·  "
        "Tom celle = ingen salg det år  ·  Antall = totalt solgte enheter",
        fmt["note"])


# ──────────────────────────────────────────────────────────────────────────────
# Ark 7 – Data  (offisiell Excel-tabell med autofilter)
# ──────────────────────────────────────────────────────────────────────────────
def _build_data(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(DARK_GREY)
    ws.set_zoom(90)
    _page_setup(ws, orientation="landscape")

    df = data["df"]

    col_info = [
        ("date",         "Dato",                    14),
        ("year",         "År",                       8),
        ("month",        "Måned",                    8),
        ("quarter",      "Kvartal",                  9),
        ("brand",        "Varemerke / Sektor",       26),
        ("units",        "Antall",                  11),
        ("net_sales",    "Nettoomsetning (NOK)",     18),
        ("article_no",   "Art.nr.",                  20),
        ("article_desc", "Artikkelbeskrivelse",      36),
    ]

    for ci, (_, _, w) in enumerate(col_info):
        ws.set_column(ci, ci, w)

    # Tittelrad
    ws.set_row(0, 34)
    ws.merge_range(0, 0, 0, len(col_info) - 1,
                   f"  {report_name}  |  Data  —  Rensede transaksjonsdata",
                   fmt["title"])

    ws.freeze_panes(2, 0)

    # Bygg tabelldata (list of lists for add_table)
    money_cols = {"net_sales"}
    int_cols   = {"units", "year", "month", "quarter"}

    table_data = []
    for _, row in df.iterrows():
        row_vals = []
        for col, _, _ in col_info:
            val = row.get(col, "")
            if val == "" or (isinstance(val, float) and math.isnan(val)):
                row_vals.append("")
            else:
                row_vals.append(val)
        table_data.append(row_vals)

    # Legg til som offisiell Excel-tabell (ListObject) med autofilter og strippede rader
    ws.add_table(
        1, 0,
        1 + len(df), len(col_info) - 1,
        {
            "name":       "SalgsData",
            "style":      "Table Style Medium 2",
            "autofilter": True,
            "first_column": False,
            "columns":    [{"header": hdr} for _, hdr, _ in col_info],
            "data":       table_data,
        }
    )

    # Kondisjonell formatering på Nettoomsetning
    if len(df) > 0:
        nok_col_idx = next(i for i, (c, _, _) in enumerate(col_info) if c == "net_sales")
        ws.conditional_format(
            2, nok_col_idx, 1 + len(df), nok_col_idx,
            {"type": "data_bar", "bar_color": MID_BLUE, "data_bar_2010": True}
        )


# ──────────────────────────────────────────────────────────────────────────────
# Offentlig inngangspunkt
# ──────────────────────────────────────────────────────────────────────────────
def generate_dashboard(data: dict, report_name: str = "Rapport") -> bytes:
    """
    Genererer nettoomsetningsrapporten med 7 regneark.
    Bruker dict fra data_processor.process().
    Returnerer .xlsx-bytes.
    """
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True, "nan_inf_to_errors": True})
    fmt = _add_formats(wb)

    ws1 = wb.add_worksheet("Dashbord")
    ws2 = wb.add_worksheet("Trendanalyse")
    ws3 = wb.add_worksheet("Varemerkeanalyse")
    ws4 = wb.add_worksheet("XYZ-analyse")
    ws5 = wb.add_worksheet("Portefølje")
    ws6 = wb.add_worksheet("Topp-artikler")
    ws7 = wb.add_worksheet("Data")

    _build_dashbord(wb,         ws1, data, fmt, report_name)
    _build_trendanalyse(wb,     ws2, data, fmt, report_name)
    _build_varemerkeanalyse(wb, ws3, data, fmt, report_name)
    _build_xyz_analyse(wb,      ws4, data, fmt, report_name)
    _build_portfolje(wb,        ws5, data, fmt, report_name)
    _build_topp_artikler(wb,    ws6, data, fmt, report_name)
    _build_data(wb,             ws7, data, fmt, report_name)

    ws1.activate()
    wb.close()
    output.seek(0)
    return output.read()
