"""
excel_generator.py
Genererer nettoomsetningsrapporten med 7 regneark.
Bruker dict fra data_processor.process().
"""
import io
import math
from datetime import datetime

import xlsxwriter

# ── Måneder ────────────────────────────────────────────────────────────────
MONTHS_EN = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
MONTHS_NO = ["Jan","Feb","Mar","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Des"]
_M_MAP = dict(zip(MONTHS_EN, MONTHS_NO))   # intern nøkkel → visningsnavn

# ── Farger ─────────────────────────────────────────────────────────────────
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


def _v(val):
    """Returner val, eller '' hvis None/NaN."""
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
    return f"NOK {v:,.0f}"


def _add_formats(wb):
    f = {}

    # ── Tittel / banner ───────────────────────────────────────────────────
    f["title"] = wb.add_format({
        "bold": True, "font_size": 18, "font_color": WHITE,
        "bg_color": DARK_BLUE, "align": "left", "valign": "vcenter",
    })
    f["subtitle"] = wb.add_format({
        "font_size": 9, "font_color": "#AACCE8", "italic": True,
        "bg_color": DARK_BLUE, "align": "left", "valign": "vcenter",
    })
    f["section_hdr"] = wb.add_format({
        "bold": True, "font_size": 10, "font_color": WHITE,
        "bg_color": MID_BLUE, "align": "left", "valign": "vcenter",
        "left": 3, "left_color": GOLD_BORDER,
    })

    # ── KPI-fliser ────────────────────────────────────────────────────────
    f["kpi_label"] = wb.add_format({
        "bold": True, "font_size": 9, "font_color": DARK_GREY,
        "bg_color": LIGHT_BLUE, "align": "center", "valign": "bottom",
        "top": 2, "top_color": MID_BLUE,
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_value"] = wb.add_format({
        "bold": True, "font_size": 18, "font_color": DARK_BLUE,
        "bg_color": LIGHT_BLUE, "align": "center", "valign": "vcenter",
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_value_nok"] = wb.add_format({
        "bold": True, "font_size": 16, "font_color": DARK_BLUE,
        "bg_color": LIGHT_BLUE, "align": "center", "valign": "vcenter",
        "num_format": '#,##0 "NOK"',
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_value_pct"] = wb.add_format({
        "bold": True, "font_size": 18, "font_color": DARK_BLUE,
        "bg_color": LIGHT_BLUE, "align": "center", "valign": "vcenter",
        "num_format": "0.0%",
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_text"] = wb.add_format({
        "bold": True, "font_size": 13, "font_color": DARK_BLUE,
        "bg_color": LIGHT_BLUE, "align": "center", "valign": "vcenter",
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_delta_pos"] = wb.add_format({
        "font_size": 9, "font_color": ACCENT_GREEN,
        "bg_color": LIGHT_BLUE, "align": "center", "valign": "top",
        "num_format": '+0.0%;-0.0%',
        "bottom": 2, "bottom_color": MID_BLUE,
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_delta_neg"] = wb.add_format({
        "font_size": 9, "font_color": ACCENT_RED,
        "bg_color": LIGHT_BLUE, "align": "center", "valign": "top",
        "num_format": '+0.0%;-0.0%',
        "bottom": 2, "bottom_color": MID_BLUE,
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })
    f["kpi_delta_na"] = wb.add_format({
        "font_size": 9, "font_color": DARK_GREY,
        "bg_color": LIGHT_BLUE, "align": "center", "valign": "top",
        "bottom": 2, "bottom_color": MID_BLUE,
        "left": 1, "left_color": MID_BLUE, "right": 1, "right_color": MID_BLUE,
    })

    # ── Kolonneoverskrifter ────────────────────────────────────────────────
    f["col_hdr"] = wb.add_format({
        "bold": True, "font_size": 10, "font_color": WHITE,
        "bg_color": DARK_BLUE, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": MID_BLUE, "text_wrap": True,
    })
    f["col_hdr_left"] = wb.add_format({
        "bold": True, "font_size": 10, "font_color": WHITE,
        "bg_color": DARK_BLUE, "align": "left", "valign": "vcenter",
        "border": 1, "border_color": MID_BLUE,
    })

    # ── Tabellceller ──────────────────────────────────────────────────────
    def _cell(bold=False, align="left", num_format=None, bg=None):
        d = {
            "font_size": 10, "align": align, "valign": "vcenter",
            "border": 1, "border_color": MID_GREY,
        }
        if bold:       d["bold"] = True
        if num_format: d["num_format"] = num_format
        if bg:         d["bg_color"] = bg
        return wb.add_format(d)

    f["cell"]             = _cell()
    f["cell_c"]           = _cell(align="center")
    f["cell_nok"]         = _cell(align="right",  num_format='#,##0 "NOK"')
    f["cell_pct"]         = _cell(align="center", num_format="0.0%")
    f["cell_pct_yoy"]     = _cell(align="center", num_format='+0.0%;-0.0%;"-"')
    f["cell_int"]         = _cell(align="center", num_format="#,##0")
    f["cell_2dec"]        = _cell(align="center", num_format="0.00")

    f["cell_a"]           = _cell(bg=LIGHT_GREY)
    f["cell_c_a"]         = _cell(align="center", bg=LIGHT_GREY)
    f["cell_nok_a"]       = _cell(align="right",  num_format='#,##0 "NOK"', bg=LIGHT_GREY)
    f["cell_pct_a"]       = _cell(align="center", num_format="0.0%",        bg=LIGHT_GREY)
    f["cell_pct_yoy_a"]   = _cell(align="center", num_format='+0.0%;-0.0%;"-"', bg=LIGHT_GREY)
    f["cell_int_a"]       = _cell(align="center", num_format="#,##0",       bg=LIGHT_GREY)
    f["cell_2dec_a"]      = _cell(align="center", num_format="0.00",        bg=LIGHT_GREY)

    # Totalrader
    f["total"]         = _cell(bold=True, bg=LIGHT_BLUE)
    f["total_c"]       = _cell(bold=True, align="center", bg=LIGHT_BLUE)
    f["total_nok"]     = _cell(bold=True, align="right",  num_format='#,##0 "NOK"', bg=LIGHT_BLUE)
    f["total_pct"]     = _cell(bold=True, align="center", num_format="0.0%",        bg=LIGHT_BLUE)
    f["total_pct_yoy"] = _cell(bold=True, align="center", num_format='+0.0%;-0.0%;"-"', bg=LIGHT_BLUE)

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
    f["abc_A"] = wb.add_format({"bold": True, "font_size": 10, "align": "center",
                                "valign": "vcenter", "font_color": WHITE,
                                "bg_color": ACCENT_GREEN, "border": 1, "border_color": MID_GREY})
    f["abc_B"] = wb.add_format({"bold": True, "font_size": 10, "align": "center",
                                "valign": "vcenter", "font_color": WHITE,
                                "bg_color": MID_BLUE, "border": 1, "border_color": MID_GREY})
    f["abc_C"] = wb.add_format({"bold": True, "font_size": 10, "align": "center",
                                "valign": "vcenter", "font_color": WHITE,
                                "bg_color": DARK_GREY, "border": 1, "border_color": MID_GREY})

    # XYZ-merker
    f["xyz_X"] = wb.add_format({"bold": True, "font_size": 10, "align": "center",
                                "valign": "vcenter", "font_color": WHITE,
                                "bg_color": TEAL, "border": 1, "border_color": MID_GREY})
    f["xyz_Y"] = wb.add_format({"bold": True, "font_size": 10, "align": "center",
                                "valign": "vcenter", "font_color": DARK_BLUE,
                                "bg_color": GOLD_BG, "border": 1, "border_color": MID_GREY})
    f["xyz_Z"] = wb.add_format({"bold": True, "font_size": 10, "align": "center",
                                "valign": "vcenter", "font_color": WHITE,
                                "bg_color": ACCENT_RED, "border": 1, "border_color": MID_GREY})

    # Portfolio-kategorier
    f["cat_stjerne"]  = wb.add_format({"bold": True, "font_size": 10, "align": "center",
                                       "valign": "vcenter", "font_color": DARK_BLUE,
                                       "bg_color": GOLD_BG, "border": 1, "border_color": MID_GREY})
    f["cat_melkeku"]  = wb.add_format({"bold": True, "font_size": 10, "align": "center",
                                       "valign": "vcenter", "font_color": WHITE,
                                       "bg_color": ACCENT_GREEN, "border": 1, "border_color": MID_GREY})
    f["cat_spm"]      = wb.add_format({"bold": True, "font_size": 10, "align": "center",
                                       "valign": "vcenter", "font_color": WHITE,
                                       "bg_color": MID_BLUE, "border": 1, "border_color": MID_GREY})
    f["cat_hund"]     = wb.add_format({"bold": True, "font_size": 10, "align": "center",
                                       "valign": "vcenter", "font_color": WHITE,
                                       "bg_color": DARK_GREY, "border": 1, "border_color": MID_GREY})
    f["cat_ny"]       = wb.add_format({"bold": True, "font_size": 10, "align": "center",
                                       "valign": "vcenter", "font_color": WHITE,
                                       "bg_color": ACCENT_ORG, "border": 1, "border_color": MID_GREY})

    # Rangmerker
    f["rank_gold"] = wb.add_format({
        "bold": True, "font_size": 10, "font_color": DARK_BLUE,
        "bg_color": GOLD_BG, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": MID_GREY,
    })
    f["rank_std"] = wb.add_format({
        "bold": True, "font_size": 10, "font_color": WHITE,
        "bg_color": MID_BLUE, "align": "center", "valign": "vcenter",
        "border": 1, "border_color": MID_GREY,
    })

    f["bullet"]     = wb.add_format({"font_size": 10, "align": "left", "valign": "vcenter", "indent": 1})
    f["blank"]      = wb.add_format({"bg_color": WHITE})
    f["blank_dark"] = wb.add_format({"bg_color": DARK_BLUE})
    f["note"]       = wb.add_format({
        "italic": True, "font_size": 8, "font_color": DARK_GREY,
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    f["insight_bold"] = wb.add_format({
        "bold": True, "font_size": 10, "align": "left", "valign": "vcenter",
        "indent": 1, "font_color": DARK_BLUE,
    })
    f["insight_body"] = wb.add_format({
        "font_size": 10, "align": "left", "valign": "vcenter",
        "indent": 1, "font_color": "#1A1A2E",
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


# ──────────────────────────────────────────────────────────────────────────────
# Ark 1 – Dashbord
# ──────────────────────────────────────────────────────────────────────────────
def _build_dashbord(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(DARK_BLUE)
    ws.set_zoom(90)

    import pandas as pd

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

    ws.set_column(0, 0, 2)
    for c in range(1, 15):
        ws.set_column(c, c, 12)
    ws.set_column(14, 18, 8)

    ws.set_row(0, 32); ws.set_row(1, 16); ws.set_row(2, 14)
    ws.merge_range(0, 0, 0, 18, f"  {report_name}  |  Nettoomsetningsrapport", fmt["title"])
    ws.merge_range(1, 0, 1, 18,
        f"  Ledersammendrag  ·  {min(all_years)}–{max_year}  ·  Alle tall i NOK",
        fmt["subtitle"])
    ws.merge_range(2, 0, 2, 18,
        f"  Generert: {datetime.now().strftime('%d.%m.%Y  %H:%M')}",
        fmt["subtitle"])

    ws.set_row(3, 6); ws.set_row(4, 18)
    ws.merge_range(4, 0, 4, 18, "  NØKKELTALL (KPI)", fmt["section_hdr"])

    ws.set_row(5, 4); ws.set_row(6, 18); ws.set_row(7, 38); ws.set_row(8, 16)

    prev_year = max_year - 1
    ytd_yoy = None
    if fy_vals.get(prev_year, 0) != 0:
        ytd_yoy = (ytd_vals.get(max_year, 0) - ytd_vals.get(prev_year, 0)) / ytd_vals[prev_year]
    fy_yoy = None
    if fy_vals.get(prev_year, 0) != 0:
        fy_yoy = (fy_vals.get(max_year, 0) - fy_vals[prev_year]) / fy_vals[prev_year]

    kr = 6
    _write_kpi_tile(ws, fmt, kr,  1, "TOTAL NETTOOMSETNING",    grand_total,                     "kpi_value_nok")
    _write_kpi_tile(ws, fmt, kr,  3, f"FY {max_year}",          fy_vals.get(max_year, 0),        "kpi_value_nok",  fy_yoy)
    _write_kpi_tile(ws, fmt, kr,  5, f"Hittil i år {ytd_label}", ytd_vals.get(max_year, 0),      "kpi_value_nok",  ytd_yoy)
    _write_kpi_tile(ws, fmt, kr,  7, "CAGR",                    cagr if cagr is not None else 0, "kpi_value_pct")
    _write_kpi_tile(ws, fmt, kr,  9, "TOPP-VAREMERKE",          top1_brand,                      "kpi_text")
    _write_kpi_tile(ws, fmt, kr, 11, "TOPP-VAREMERKE ANDEL",    top1_share,                      "kpi_value_pct")
    _write_kpi_tile(ws, fmt, kr, 13, "TOPP 3 ANDEL",            top3_share,                      "kpi_value_pct")

    # ── Ledersammendrag ──────────────────────────────────────────────────
    ws.set_row(9, 6); ws.set_row(10, 18)
    ws.merge_range(10, 0, 10, 18, "  LEDERSAMMENDRAG", fmt["section_hdr"])

    bullets = []
    bullets.append(
        f"▸  Total nettoomsetning alle år: {_nok(grand_total)}  |  "
        f"{len(all_years)} år med data ({min(all_years)}–{max_year})"
    )
    if fy_yoy is not None:
        arrow = "▲" if fy_yoy >= 0 else "▼"
        bullets.append(
            f"▸  FY {max_year}: {_nok(fy_vals.get(max_year, 0))}  "
            f"{arrow} {abs(fy_yoy)*100:.1f}% sammenlignet med FY {prev_year}"
        )
    else:
        bullets.append(f"▸  FY {max_year}: {_nok(fy_vals.get(max_year, 0))}")

    if ytd_yoy is not None:
        arrow = "▲" if ytd_yoy >= 0 else "▼"
        bullets.append(
            f"▸  Hittil i år ({ytd_label}) {max_year}: {_nok(ytd_vals.get(max_year, 0))}  "
            f"{arrow} {abs(ytd_yoy)*100:.1f}% vs {ytd_label} {prev_year}"
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
    if not ann_sum.empty and "net_sales" in ann_sum.columns:
        ns = pd.to_numeric(ann_sum["net_sales"], errors="coerce")
        valid = ann_sum[ns.notna()]
        if len(valid):
            best = valid.loc[ns[ns.notna()].idxmax()]
            bullets.append(f"▸  Beste år etter omsetning: {best['year']}  ({_nok(float(best['net_sales']))})")

    # Gini-tillegg
    gini_lbl = ("Svært skjev" if gini > 0.6 else "Moderat skjev" if gini > 0.35 else "Relativt jevn")
    bullets.append(
        f"▸  Gini-koeffisient (omsetningsfordeling): {gini:.3f}  —  {gini_lbl} fordeling mellom varemerker"
    )

    ws.set_row(11, 4)
    for i, b in enumerate(bullets):
        ws.set_row(12 + i, 16)
        ws.merge_range(12 + i, 1, 12 + i, 18, b, fmt["bullet"])

    last_bul = 12 + len(bullets) - 1

    # ── Årssammendrag ────────────────────────────────────────────────────
    ws.set_row(last_bul + 1, 6)
    tbl = last_bul + 2
    ws.set_row(tbl, 18)
    ws.merge_range(tbl, 0, tbl, 18, "  ÅRSSAMMENDRAG NETTOOMSETNING", fmt["section_hdr"])
    tbl += 1

    ws.set_row(tbl, 18)
    ann_hdrs = ["År", "Nettoomsetning (NOK)", "ÅoÅ-vekst",
                f"Hittil i år {ytd_label} (NOK)", "Hittil YoY", "Beste kvartal"]
    ann_col_w = [8, 20, 13, 22, 13, 14]
    for ci, (h, w) in enumerate(zip(ann_hdrs, ann_col_w)):
        hf = fmt["col_hdr_left"] if ci == 0 else fmt["col_hdr"]
        ws.write(tbl, 1 + ci, h, hf)
        ws.set_column(1 + ci, 1 + ci, w)
    tbl += 1

    for ri, (_, row) in enumerate(ann_sum.iterrows()):
        is_alt = ri % 2 == 1
        lf = fmt["cell_a"] if is_alt else fmt["cell"]
        nf = fmt["cell_nok_a"] if is_alt else fmt["cell_nok"]
        pf = fmt["cell_pct_yoy_a"] if is_alt else fmt["cell_pct_yoy"]
        cf = fmt["cell_c_a"] if is_alt else fmt["cell_c"]
        ws.set_row(tbl + ri, 16)
        ws.write(tbl + ri, 1, str(row.get("year", "")), lf)
        ws.write(tbl + ri, 2, _v(row.get("net_sales", 0)), nf)
        yoy = _pct(row.get("yoy"))
        ws.write(tbl + ri, 3, yoy, pf if yoy != "" else lf)
        ws.write(tbl + ri, 4, _v(row.get("ytd", 0)), nf)
        ytd_y = _pct(row.get("ytd_yoy"))
        ws.write(tbl + ri, 5, ytd_y, pf if ytd_y != "" else lf)
        ws.write(tbl + ri, 6, str(row.get("best_q", "—")), cf)

    # ── Strategiske innsikter ────────────────────────────────────────────
    pareto_df = data.get("pareto", pd.DataFrame())
    n_brands_80 = 0
    if not pareto_df.empty and "cumulative" in pareto_df.columns:
        n_brands_80 = int((pareto_df["cumulative"] <= 0.80).sum()) + 1

    n_a = sum(1 for v in abc_brands.values() if v == "A")
    n_b = sum(1 for v in abc_brands.values() if v == "B")
    n_c = sum(1 for v in abc_brands.values() if v == "C")

    hhi_lbl = ("Høy konsentrasjon" if hhi > 0.25 else
               "Moderat konsentrasjon" if hhi > 0.15 else
               "Diversifisert")
    peak_idx   = seas.get(peak_m, 1.0)
    trough_idx = seas.get(trough_m, 1.0)

    si_row = tbl + len(ann_sum) + 2
    ws.set_row(si_row, 18)
    ws.merge_range(si_row, 0, si_row, 18, "  STRATEGISKE INNSIKTER", fmt["section_hdr"])
    si_row += 1; ws.set_row(si_row, 6); si_row += 1

    insights = [
        ("PORTEFØLJEKONSENTRASJON  (Herfindahl-Hirschman-indeks)",
         f"HHI = {hhi:.3f}  →  {hhi_lbl}.  "
         + ("Høy avhengighet av få varemerker gir omsetningsrisiko — vurder å bredde porteføljen."
            if hhi > 0.25 else
            "Moderat spredning gir rimelig motstandskraft."
            if hhi > 0.15 else
            "Omsetningen er godt fordelt — lav enkeltmerkeavhengighet.")),
        ("ABC-KLASSIFISERING  (basert på siste hele år)",
         f"A-klasse: {n_a} varemerker (~70% av omsetning)  ·  "
         f"B-klasse: {n_b} (neste 20%)  ·  "
         f"C-klasse: {n_c} (bunn 10%)  —  "
         "Prioriter ressurser på A-klasse; evaluer C-klasse for rasjonalisering."),
        ("PARETO-ANALYSE  (80/20-regelen)",
         f"{n_brands_80} varemerke(r) står for 80% av total omsetning  "
         f"(av {len(abc_brands)} totalt).  "
         + ("Sterk Pareto-konsentrasjon — avhengighetsrisiko er forhøyet."
            if n_brands_80 <= 3 else
            "Sunn Pareto-fordeling på tvers av varemerkeporteføljen.")),
        ("GINI-KOEFFISIENT  (omsetningsulikhet mellom varemerker)",
         f"Gini = {gini:.3f}  —  {'Svært skjev' if gini > 0.6 else 'Moderat skjev' if gini > 0.35 else 'Relativt jevn'} "
         f"fordeling.  {'Vurder tiltak for å redusere konsentrasjon.' if gini > 0.5 else 'Akseptabelt diversifisert portefølje.'}"),
        ("SESONGMØNSTER",
         f"Toppmåned: {peak_m} (indeks {peak_idx:.2f}×)  ·  "
         f"Bunnmåned: {trough_m} (indeks {trough_idx:.2f}×).  "
         f"Topp/bunn-ratio: {peak_idx / max(trough_idx, 0.01):.1f}×  —  "
         + ("Høy sesongsvingning: planlegg lager og likviditet deretter."
            if peak_idx / max(trough_idx, 0.01) > 2.0 else
            "Moderat sesongvariasjon — relativt stabil etterspørsel gjennom året.")),
    ]

    for label, detail in insights:
        ws.set_row(si_row, 15)
        ws.merge_range(si_row, 1, si_row, 18, f"  {label}", fmt["insight_bold"])
        si_row += 1
        ws.set_row(si_row, 28)
        ws.merge_range(si_row, 1, si_row, 18, f"  {detail}", fmt["insight_body"])
        si_row += 1
        ws.set_row(si_row, 4); si_row += 1

    ws.set_row(si_row, 14)
    ws.merge_range(si_row, 1, si_row, 18,
        "  Rammeverk: HHI (Herfindahl-Hirschman 1964)  ·  ABC-analyse (Dickie 1951)  "
        "·  Pareto-prinsippet  ·  Gini (1912)  ·  CAGR-vekstbenchmarking",
        fmt["note"])


# ──────────────────────────────────────────────────────────────────────────────
# Ark 2 – Trendanalyse
# ──────────────────────────────────────────────────────────────────────────────
def _build_trendanalyse(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(MID_BLUE)
    ws.set_zoom(90)

    all_years     = data["all_years"]
    monthly_pivot = data["monthly_pivot"]
    monthly_yoy   = data["monthly_yoy"]
    qpivot        = data["quarterly_pivot"]
    seas_dict     = data["seasonality"]

    ws.set_column(0, 0, 8)
    for c in range(1, 13):
        ws.set_column(c, c, 9)
    ws.set_column(13, 13, 12)
    ws.set_column(14, 16, 9)

    ws.set_row(0, 32); ws.set_row(1, 14)
    ws.merge_range(0, 0, 0, 16, f"  {report_name}  |  Trendanalyse  —  Nettoomsetning per periode", fmt["title"])
    ws.merge_range(1, 0, 1, 16, "  Alle tall i NOK", fmt["subtitle"])
    ws.set_row(2, 6)

    # ── Månedlig pivot ───────────────────────────────────────────────────
    # Data positions tracked for chart
    monthly_hdr_row  = 3   # section header
    monthly_col_row  = 4   # column header row
    monthly_data_row = 5   # first data row

    ws.set_row(monthly_hdr_row, 18)
    ws.merge_range(monthly_hdr_row, 0, monthly_hdr_row, 13,
                   "  MÅNEDLIG NETTOOMSETNING (NOK)", fmt["section_hdr"])

    ws.set_row(monthly_col_row, 18)
    ws.write(monthly_col_row, 0, "År", fmt["col_hdr_left"])
    for mi, m_en in enumerate(MONTHS_EN):
        ws.write(monthly_col_row, 1 + mi, _M_MAP.get(m_en, m_en), fmt["col_hdr"])
    ws.write(monthly_col_row, 13, "TOTALT", fmt["col_hdr"])

    n_data_years = 0  # years with data (not TOTAL)
    r = monthly_data_row
    for ri, year_str in enumerate(monthly_pivot.index):
        is_total = year_str == "TOTAL"
        is_alt   = ri % 2 == 1 and not is_total
        lf = fmt["total"]     if is_total else (fmt["cell_a"]     if is_alt else fmt["cell"])
        nf = fmt["total_nok"] if is_total else (fmt["cell_nok_a"] if is_alt else fmt["cell_nok"])
        ws.set_row(r + ri, 16)
        ws.write(r + ri, 0, year_str, lf)
        for mi, m_en in enumerate(MONTHS_EN):
            v = float(monthly_pivot.loc[year_str, m_en]) if m_en in monthly_pivot.columns else 0
            ws.write(r + ri, 1 + mi, v if v else "",
                     nf if v else (fmt["cell_a"] if is_alt else fmt["cell"]))
        tot = float(monthly_pivot.loc[year_str, "TOTAL"]) if "TOTAL" in monthly_pivot.columns else 0
        ws.write(r + ri, 13, tot, nf)
        if not is_total:
            n_data_years += 1

    # Chart: linjediagram månedlig omsetning per år
    chart = wb.add_chart({"type": "line"})
    colors = [MID_BLUE, ACCENT_GREEN, ACCENT_RED, PURPLE, TEAL, ACCENT_ORG, DARK_BLUE]
    for i in range(n_data_years):
        chart.add_series({
            "name":       ["Trendanalyse", monthly_col_row, 0, monthly_col_row, 0] if False else str(monthly_pivot.index[i]),
            "categories": ["Trendanalyse", monthly_col_row, 1, monthly_col_row, 12],
            "values":     ["Trendanalyse", monthly_data_row + i, 1, monthly_data_row + i, 12],
            "line":       {"width": 2.5, "color": colors[i % len(colors)]},
            "marker":     {"type": "circle", "size": 5,
                           "fill": {"color": colors[i % len(colors)]},
                           "border": {"color": WHITE}},
        })
    chart.set_title({"name": "Månedlig nettoomsetning per år (NOK)"})
    chart.set_x_axis({"name": "Måned"})
    chart.set_y_axis({"name": "NOK", "num_format": "#,##0"})
    chart.set_legend({"position": "bottom"})
    chart.set_style(10)
    chart.set_size({"width": 680, "height": 320})

    next_r = r + len(monthly_pivot) + 2

    ws.insert_chart(next_r, 0, chart, {"x_offset": 2, "y_offset": 4})
    chart_rows = 18  # rows occupied by chart
    next_r += chart_rows

    # ── Månedlig YoY-vekst ───────────────────────────────────────────────
    ws.set_row(next_r, 18)
    ws.merge_range(next_r, 0, next_r, 13, "  MÅNEDLIG ÅOÅ-VEKST %", fmt["section_hdr"])
    next_r += 1
    ws.set_row(next_r, 18)
    ws.write(next_r, 0, "Periode", fmt["col_hdr_left"])
    for mi, m_en in enumerate(MONTHS_EN):
        ws.write(next_r, 1 + mi, _M_MAP.get(m_en, m_en), fmt["col_hdr"])
    ws.write(next_r, 13, "Hele år", fmt["col_hdr"])
    next_r += 1

    for ri, (_, row) in enumerate(monthly_yoy.iterrows()):
        is_alt = ri % 2 == 1
        lf = fmt["cell_a"] if is_alt else fmt["cell"]
        pf = fmt["cell_pct_yoy_a"] if is_alt else fmt["cell_pct_yoy"]
        ws.set_row(next_r + ri, 16)
        ws.write(next_r + ri, 0, str(row.get("label", "")), lf)
        for mi, m_en in enumerate(MONTHS_EN):
            v = _pct(row.get(m_en))
            ws.write(next_r + ri, 1 + mi, v, pf if v != "" else lf)
        fy_v = _pct(row.get("Full Year"))
        ws.write(next_r + ri, 13, fy_v, pf if fy_v != "" else lf)

    next_r += max(len(monthly_yoy), 1) + 2

    # ── Kvartalsoversikt ─────────────────────────────────────────────────
    ws.set_row(next_r, 18)
    ws.merge_range(next_r, 0, next_r, 6, "  KVARTALSVIS NETTOOMSETNING (NOK)", fmt["section_hdr"])
    next_r += 1
    ws.set_row(next_r, 18)
    for ci, h in enumerate(["År", "K1", "K2", "K3", "K4", "Totalt", "K2-andel"]):
        hf = fmt["col_hdr_left"] if ci == 0 else fmt["col_hdr"]
        ws.write(next_r, ci, h, hf)
    next_r += 1

    for ri, (_, row) in enumerate(qpivot.iterrows()):
        is_alt = ri % 2 == 1
        lf = fmt["cell_a"] if is_alt else fmt["cell"]
        nf = fmt["cell_nok_a"] if is_alt else fmt["cell_nok"]
        pf = fmt["cell_pct_a"] if is_alt else fmt["cell_pct"]
        ws.set_row(next_r + ri, 16)
        ws.write(next_r + ri, 0, str(row.get("Year", "")), lf)
        for qi, q in enumerate(["Q1", "Q2", "Q3", "Q4"]):
            v = row.get(q, 0)
            ws.write(next_r + ri, 1 + qi, v if v else "", nf if v else lf)
        tot = row.get("Total", 0)
        ws.write(next_r + ri, 5, tot if tot else "", nf if tot else lf)
        q2s = _pct(row.get("Q2 Share"))
        ws.write(next_r + ri, 6, q2s, pf if q2s != "" else lf)

    next_r += len(qpivot) + 2

    # ── Sesongindeks ─────────────────────────────────────────────────────
    ws.set_row(next_r, 18)
    ws.merge_range(next_r, 0, next_r, 13,
                   "  SESONGINDEKS  (1,00 = gjennomsnittsmåned, kun hele kalenderår)",
                   fmt["section_hdr"])
    next_r += 1
    ws.set_row(next_r, 18)
    ws.write(next_r, 0, "Måned", fmt["col_hdr_left"])
    for mi, m_en in enumerate(MONTHS_EN):
        ws.write(next_r, 1 + mi, _M_MAP.get(m_en, m_en), fmt["col_hdr"])
    next_r += 1
    ws.set_row(next_r, 18)
    ws.write(next_r, 0, "Indeks", fmt["cell"])
    for mi, m_en in enumerate(MONTHS_EN):
        v = seas_dict.get(m_en, 1.0)
        sf = fmt["seas_hi"] if v >= 1.1 else (fmt["seas_lo"] if v <= 0.9 else fmt["seas_mid"])
        ws.write(next_r, 1 + mi, v, sf)
    next_r += 2
    ws.merge_range(next_r, 0, next_r, 13,
        "  Grønt = sesongtopp (≥ 1,10)  ·  Rødt = sesongbunn (≤ 0,90)  ·  "
        "Beregnet som månedlig gjennomsnitt / totalt gjennomsnitt over hele år",
        fmt["note"])


# ──────────────────────────────────────────────────────────────────────────────
# Ark 3 – Varemerkeanalyse
# ──────────────────────────────────────────────────────────────────────────────
def _build_varemerkeanalyse(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(ACCENT_GREEN)
    ws.set_zoom(90)

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

    ws.set_column(0, 0, 24)
    for c in range(1, 1 + n_fy):
        ws.set_column(c, c, 14)
    ws.set_column(1 + n_fy, 1 + n_fy, 14)
    for c in range(2 + n_fy, 2 + n_fy + n_yoy):
        ws.set_column(c, c, 11)
    ws.set_column(2 + n_fy + n_yoy, 2 + n_fy + n_yoy, 11)
    ws.set_column(3 + n_fy + n_yoy, 3 + n_fy + n_yoy, 7)

    ws.set_row(0, 32); ws.set_row(1, 14)
    ws.merge_range(0, 0, 0, n_cols,
                   f"  {report_name}  |  Varemerkeanalyse  —  Nettoomsetning per varemerke",
                   fmt["title"])
    ws.merge_range(1, 0, 1, n_cols,
                   "  Alle tall i NOK  ·  Sortert etter siste hele år  ·  "
                   "ABC = porteføljeklassifisering (Dickie 1951)",
                   fmt["subtitle"])
    ws.set_row(2, 6)

    ws.set_row(3, 18)
    ws.merge_range(3, 0, 3, n_cols, "  VAREMERKEANALYSE", fmt["section_hdr"])

    ws.set_row(4, 18)
    headers = (["Varemerke"] + [f"FY {y}" for y in all_years] +
               [f"Hittil i år {ytd_label}"] +
               [f"ÅoÅ {all_years[i]}v{all_years[i-1]}" for i in range(1, len(all_years))] +
               [f"Andel FY{last_full}", "ABC"])
    for ci, h in enumerate(headers):
        hf = fmt["col_hdr_left"] if ci == 0 else fmt["col_hdr"]
        ws.write(4, ci, h, hf)

    for ri, (_, row) in enumerate(brand_perf.iterrows()):
        brand_name = str(row.get("brand", ""))
        is_total   = brand_name.upper() == "TOTAL"
        is_alt     = ri % 2 == 1 and not is_total
        lf  = fmt["total"]         if is_total else (fmt["cell_a"]         if is_alt else fmt["cell"])
        nf  = fmt["total_nok"]     if is_total else (fmt["cell_nok_a"]     if is_alt else fmt["cell_nok"])
        pf  = fmt["total_pct_yoy"] if is_total else (fmt["cell_pct_yoy_a"] if is_alt else fmt["cell_pct_yoy"])
        sf  = fmt["total_pct"]     if is_total else (fmt["cell_pct_a"]     if is_alt else fmt["cell_pct"])
        ws.set_row(5 + ri, 16)

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
            ws.write(5 + ri, ci, yoy_v if yoy_v != "" else "-",
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

    pareto_section = 5 + len(brand_perf) + 3

    # ── Pareto-analyse ───────────────────────────────────────────────────
    ws.set_row(pareto_section, 18)
    ws.merge_range(pareto_section, 0, pareto_section, 5,
                   "  PARETO-ANALYSE  —  80/20 Omsetningskonsentrasjon", fmt["section_hdr"])
    pareto_section += 1
    ws.set_row(pareto_section, 18)
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
        ws.set_row(pareto_section + ri, 16)
        ws.write(pareto_section + ri, 0, int(row.get("rank", ri+1)),         lf)
        ws.write(pareto_section + ri, 1, str(row.get("brand", "")),          bf)
        ws.write(pareto_section + ri, 2, float(row.get("total_revenue", 0)), nf)
        ws.write(pareto_section + ri, 3, float(row.get("pct", 0)),           pf)
        ws.write(pareto_section + ri, 4, float(row.get("cumulative", 0)),    pf)
        ws.write(pareto_section + ri, 5, str(row.get("threshold_80", "")),   tf)

    # Chart: topp-10 varemerker — horisontalt stolpediagram
    n_chart = min(10, len(pareto_df))
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
    bar_chart.set_size({"width": 560, "height": 380})

    chart_row = pareto_section + len(pareto_df) + 2
    ws.insert_chart(chart_row, 0, bar_chart, {"x_offset": 2, "y_offset": 4})


# ──────────────────────────────────────────────────────────────────────────────
# Ark 4 – XYZ-analyse
# ──────────────────────────────────────────────────────────────────────────────
def _build_xyz_analyse(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(TEAL)
    ws.set_zoom(90)

    xyz_df         = data.get("xyz_df")
    abc_xyz_matrix = data.get("abc_xyz_matrix", {})
    abc_brands     = data.get("abc_brands", {})

    import pandas as pd
    if xyz_df is None or (hasattr(xyz_df, "empty") and xyz_df.empty):
        ws.merge_range(0, 0, 0, 6, f"  {report_name}  |  XYZ-analyse — Ikke nok data", fmt["title"])
        return

    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 20)
    ws.set_column(2, 2, 20)
    ws.set_column(3, 3, 12)
    ws.set_column(4, 4, 12)
    ws.set_column(5, 5, 8)
    ws.set_column(6, 6, 8)
    ws.set_column(7, 7, 10)

    ws.set_row(0, 32); ws.set_row(1, 14)
    ws.merge_range(0, 0, 0, 7,
                   f"  {report_name}  |  XYZ-analyse  —  Etterspørselsvariabilitet per varemerke",
                   fmt["title"])
    ws.merge_range(1, 0, 1, 7,
                   "  Variasjonskoeffisient (CV = std/gjennomsnitt) av månedlig omsetning per varemerke  ·  "
                   "Ref: Silver, Pyke & Peterson (1998)",
                   fmt["subtitle"])
    ws.set_row(2, 6)

    # ── Forklaring ───────────────────────────────────────────────────────
    ws.set_row(3, 18)
    ws.merge_range(3, 0, 3, 7, "  METODIKK OG KLASSIFISERING", fmt["section_hdr"])
    exp_rows = [
        "  X = Stabil etterspørsel            (CV < 0,50)  —  Lav variabilitet, forutsigbar omsetning",
        "  Y = Variabel etterspørsel           (0,50 ≤ CV < 1,00)  —  Moderat variabilitet, sesong- eller trendpåvirket",
        "  Z = Svært variabel etterspørsel     (CV ≥ 1,00)  —  Høy variabilitet, vanskelig å forutsi",
    ]
    xyz_colors = [TEAL, "#B7950B", ACCENT_RED]
    for i, (txt, col) in enumerate(zip(exp_rows, xyz_colors)):
        exp_fmt = wb.add_format({
            "font_size": 10, "font_color": col, "bold": True,
            "align": "left", "valign": "vcenter", "indent": 1,
        })
        ws.set_row(4 + i, 16)
        ws.merge_range(4 + i, 0, 4 + i, 7, txt, exp_fmt)
    ws.set_row(7, 6)

    # ── XYZ-tabell ───────────────────────────────────────────────────────
    ws.set_row(8, 18)
    ws.merge_range(8, 0, 8, 7, "  XYZ-KLASSIFISERING PER VAREMERKE", fmt["section_hdr"])
    ws.set_row(9, 18)
    hdrs = ["Varemerke", "Gj.snitt månedlig (NOK)", "Std.avvik (NOK)",
            "Var.koeff. (CV)", "Antall måneder", "XYZ", "ABC", "Kombinert"]
    for ci, h in enumerate(hdrs):
        hf = fmt["col_hdr_left"] if ci == 0 else fmt["col_hdr"]
        ws.write(9, ci, h, hf)

    xyz_fmt_map = {"X": "xyz_X", "Y": "xyz_Y", "Z": "xyz_Z"}

    for ri, (_, row) in enumerate(xyz_df.iterrows()):
        is_alt = ri % 2 == 1
        lf   = fmt["cell_a"] if is_alt else fmt["cell"]
        nf   = fmt["cell_nok_a"] if is_alt else fmt["cell_nok"]
        cf   = fmt["cell_2dec_a"] if is_alt else fmt["cell_2dec"]
        intf = fmt["cell_int_a"] if is_alt else fmt["cell_int"]
        ws.set_row(10 + ri, 16)
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
        ws.write(10 + ri, 7, str(row.get("combined", "—")),     lf)

    matrix_row = 10 + len(xyz_df) + 3

    # ── ABC–XYZ-matrise ──────────────────────────────────────────────────
    ws.set_row(matrix_row, 18)
    ws.merge_range(matrix_row, 0, matrix_row, 4,
                   "  ABC–XYZ KOMBINASJONSMATRISE  (antall varemerker per kvadrant)",
                   fmt["section_hdr"])
    matrix_row += 1

    hdr_fmt = wb.add_format({"bold": True, "font_size": 11, "align": "center",
                             "valign": "vcenter", "border": 2,
                             "border_color": DARK_BLUE, "bg_color": DARK_BLUE,
                             "font_color": WHITE})
    ws.set_row(matrix_row, 22)
    ws.write(matrix_row, 0, "", hdr_fmt)
    ws.write(matrix_row, 1, "X  (stabil)", fmt["xyz_X"])
    ws.write(matrix_row, 2, "Y  (variabel)", fmt["xyz_Y"])
    ws.write(matrix_row, 3, "Z  (erratisk)", fmt["xyz_Z"])
    matrix_row += 1

    for abc_k in ["A", "B", "C"]:
        ws.set_row(matrix_row, 22)
        ws.write(matrix_row, 0, f"{abc_k}-klasse", fmt[f"abc_{abc_k}"])
        for xi, xyz_k in enumerate(["X", "Y", "Z"]):
            cnt = abc_xyz_matrix.get(abc_k, {}).get(xyz_k, 0)
            cell_fmt = wb.add_format({
                "bold": True, "font_size": 13, "align": "center",
                "valign": "vcenter", "border": 1, "border_color": MID_GREY,
                "bg_color": LIGHT_BLUE if cnt > 0 else LIGHT_GREY,
                "font_color": DARK_BLUE,
            })
            ws.write(matrix_row, 1 + xi, cnt, cell_fmt)
        matrix_row += 1

    matrix_row += 1
    ws.set_row(matrix_row, 30)
    ws.merge_range(matrix_row, 0, matrix_row, 7,
        "  AX = Høy verdi, stabil (prioriter)  ·  AZ = Høy verdi, erratisk (kritisk risiko)  ·  "
        "CX = Lav verdi, stabil (standard)  ·  CZ = Lav verdi, erratisk (vurder eliminering)  "
        "·  Ref: Scholz-Reiter et al. (2012)",
        fmt["note"])


# ──────────────────────────────────────────────────────────────────────────────
# Ark 5 – Portefølje
# ──────────────────────────────────────────────────────────────────────────────
def _build_portfolje(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(PURPLE)
    ws.set_zoom(90)

    portfolio_df  = data.get("portfolio_df")
    ref_years     = data.get("portfolio_ref_years")
    avg_growth    = data.get("portfolio_avg_growth")
    avg_share     = data.get("portfolio_avg_share")

    import pandas as pd
    has_data = (portfolio_df is not None and
                hasattr(portfolio_df, "empty") and
                not portfolio_df.empty)

    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 14)
    ws.set_column(2, 2, 14)
    ws.set_column(3, 3, 18)
    ws.set_column(4, 4, 8)
    ws.set_column(5, 8, 22)

    ws.set_row(0, 32); ws.set_row(1, 14)
    ref_str = f"{ref_years[0]}–{ref_years[1]}" if ref_years else "N/A"
    ws.merge_range(0, 0, 0, 8,
                   f"  {report_name}  |  Porteføljeanalyse  —  BCG-inspirert vekst/andel-matrise",
                   fmt["title"])
    ws.merge_range(1, 0, 1, 8,
                   f"  Referanseperiode: {ref_str}  ·  Vekstrate = ÅoÅ nettoomsetning  ·  "
                   "Andel = andel av total porteføljeomsetning  ·  Ref: Henderson (1970)",
                   fmt["subtitle"])
    ws.set_row(2, 6)

    # ── Forklaring ───────────────────────────────────────────────────────
    ws.set_row(3, 18)
    ws.merge_range(3, 0, 3, 8, "  KATEGORIER OG STRATEGISKE ANBEFALINGER", fmt["section_hdr"])

    cat_rows = [
        ("Stjerne  ★", GOLD_BG, DARK_BLUE,
         "Høy vekst + høy andel  —  Invester for å opprettholde posisjon og utnytte momentum"),
        ("Melkeku  ◆", ACCENT_GREEN, WHITE,
         "Lav vekst + høy andel  —  Optimaliser margin og bruk overskudd til å finansiere vekst"),
        ("Spørsmålstegn  ?", MID_BLUE, WHITE,
         "Høy vekst + lav andel  —  Vurder selektivt: øk investering på lovende merkevarer"),
        ("Hund  ✕", DARK_GREY, WHITE,
         "Lav vekst + lav andel  —  Evaluer for rasjonalisering eller avvikling"),
    ]
    for i, (cat_name, bg, fc, desc) in enumerate(cat_rows):
        cat_hdr_fmt = wb.add_format({"bold": True, "font_size": 10, "font_color": fc,
                                     "bg_color": bg, "align": "left", "valign": "vcenter",
                                     "indent": 1, "border": 1, "border_color": MID_GREY})
        cat_desc_fmt = wb.add_format({"font_size": 10, "font_color": "#1A1A2E",
                                      "align": "left", "valign": "vcenter", "indent": 1,
                                      "border": 1, "border_color": MID_GREY})
        ws.set_row(4 + i, 16)
        ws.write(4 + i, 0, f"  {cat_name}", cat_hdr_fmt)
        ws.merge_range(4 + i, 1, 4 + i, 8, desc, cat_desc_fmt)
    ws.set_row(8, 6)

    if not has_data:
        ws.set_row(9, 18)
        ws.merge_range(9, 0, 9, 8,
                       "  Ikke nok historiske år til porteføljeanalyse (krever minst 2 hele kalenderår)",
                       fmt["note"])
        return

    # ── Porteføljeoversikt ───────────────────────────────────────────────
    ws.set_row(9, 18)
    ws.merge_range(9, 0, 9, 8,
                   f"  PORTEFØLJEOVERSIKT  —  {ref_str}",
                   fmt["section_hdr"])
    ws.set_row(10, 18)
    hdrs = ["Varemerke", f"Omsetningsandel FY{ref_years[1] if ref_years else ''}",
            "Vekstrate ÅoÅ", "Kategori", "ABC"]
    for ci, h in enumerate(hdrs):
        hf = fmt["col_hdr_left"] if ci == 0 else fmt["col_hdr"]
        ws.write(10, ci, h, hf)

    cat_fmt_map = {
        "Stjerne":       "cat_stjerne",
        "Melkeku":       "cat_melkeku",
        "Spørsmålstegn": "cat_spm",
        "Hund":          "cat_hund",
        "Ny":            "cat_ny",
    }

    for ri, (_, row) in enumerate(portfolio_df.iterrows()):
        is_alt = ri % 2 == 1
        lf = fmt["cell_a"] if is_alt else fmt["cell"]
        pf = fmt["cell_pct_a"] if is_alt else fmt["cell_pct"]
        pyf = fmt["cell_pct_yoy_a"] if is_alt else fmt["cell_pct_yoy"]
        ws.set_row(11 + ri, 16)
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

    # Threshold note
    note_row = 11 + len(portfolio_df) + 2
    ws.set_row(note_row, 26)
    ws.merge_range(note_row, 0, note_row, 8,
        f"  Terskelverdi vekst: {avg_growth*100:.1f}% (porteføljegjennomsnitt {ref_str})  ·  "
        f"Terskelverdi andel: {avg_share*100:.2f}% (1 / antall varemerker)  ·  "
        "Verdiene er interne og relativ til denne porteføljen, ikke markedet som helhet",
        fmt["note"])

    # ── Matriseoversikt ──────────────────────────────────────────────────
    matrix_row = note_row + 3
    ws.set_row(matrix_row, 18)
    ws.merge_range(matrix_row, 0, matrix_row, 4,
                   "  MATRISEOVERSIKT  (antall varemerker per kvadrant)",
                   fmt["section_hdr"])
    matrix_row += 1

    cats = ["Stjerne", "Melkeku", "Spørsmålstegn", "Hund", "Ny"]
    counts = {c: 0 for c in cats}
    for _, row in portfolio_df.iterrows():
        cat = str(row.get("category", ""))
        if cat in counts:
            counts[cat] += 1

    ws.set_row(matrix_row, 22)
    for ci, cat in enumerate(cats):
        ws.write(matrix_row, ci, cat,
                 fmt[cat_fmt_map[cat]] if cat in cat_fmt_map else fmt["cell_c"])
    matrix_row += 1
    ws.set_row(matrix_row, 22)
    cnt_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "center",
                             "valign": "vcenter", "border": 1, "border_color": MID_GREY,
                             "bg_color": LIGHT_BLUE})
    for ci, cat in enumerate(cats):
        ws.write(matrix_row, ci, counts[cat], cnt_fmt)


# ──────────────────────────────────────────────────────────────────────────────
# Ark 6 – Topp-artikler
# ──────────────────────────────────────────────────────────────────────────────
def _build_topp_artikler(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(GOLD_BORDER)
    ws.set_zoom(90)

    top5_brands         = data.get("top5_brands", [])
    top_items_per_brand = data.get("top_items_per_brand", {})
    abc_brands          = data.get("abc_brands", {})

    ws.set_column(0, 0, 6)
    ws.set_column(1, 1, 18)
    ws.set_column(2, 2, 38)
    ws.set_column(3, 3, 12)
    ws.set_column(4, 4, 16)

    ws.set_row(0, 32); ws.set_row(1, 14)
    ws.merge_range(0, 0, 0, 4,
                   f"  {report_name}  |  Bestselgende artikler etter antall  —  Topp 5 varemerker",
                   fmt["title"])
    ws.merge_range(1, 0, 1, 4,
                   "  Rangert etter antall solgte enheter (alle tider)  ·  "
                   "Farge- og størrelsevarianter slått sammen per basisartikkel",
                   fmt["subtitle"])

    r = 2
    for brand in top5_brands:
        items = top_items_per_brand.get(brand)
        if items is None or len(items) == 0:
            continue

        abc_class  = abc_brands.get(brand, "")
        abc_suffix = f"  [{abc_class}-klasse]" if abc_class else ""

        ws.set_row(r, 6); r += 1
        ws.set_row(r, 20)
        ws.merge_range(r, 0, r, 4, f"  {brand}{abc_suffix}", fmt["section_hdr"])
        r += 1

        ws.set_row(r, 18)
        ws.write(r, 0, "Rang",         fmt["col_hdr"])
        ws.write(r, 1, "Art.nr.",       fmt["col_hdr"])
        ws.write(r, 2, "Beskrivelse",   fmt["col_hdr_left"])
        ws.write(r, 3, "Antall solgt",  fmt["col_hdr"])
        ws.write(r, 4, "Nettoomsetning (NOK)", fmt["col_hdr"])
        r += 1

        for ri, (_, item) in enumerate(items.iterrows()):
            is_alt = ri % 2 == 1
            lf  = fmt["cell_a"]     if is_alt else fmt["cell"]
            nf  = fmt["cell_nok_a"] if is_alt else fmt["cell_nok"]
            intf = fmt["cell_int_a"] if is_alt else fmt["cell_int"]
            rf  = fmt["rank_gold"] if ri == 0 else fmt["rank_std"]
            ws.set_row(r + ri, 16)
            ws.write(r + ri, 0, ri + 1,                          rf)
            ws.write(r + ri, 1, str(item.get("article_no", "")), lf)
            ws.write(r + ri, 2, str(item.get("article_desc", "")), lf)
            ws.write(r + ri, 3, int(item.get("units", 0)),       intf)
            ws.write(r + ri, 4, float(item.get("net_sales", 0)), nf)
        r += len(items)

    r += 1
    ws.set_row(r, 14)
    ws.merge_range(r, 0, r, 4,
        "  Varemerkerangering etter all-time nettoomsetning  ·  "
        "ABC basert på siste hele års omsetningsbidrag  ·  "
        "Antall = totalt antall solgte enheter på tvers av alle transaksjoner",
        fmt["note"])


# ──────────────────────────────────────────────────────────────────────────────
# Ark 7 – Data
# ──────────────────────────────────────────────────────────────────────────────
def _build_data(wb, ws, data, fmt, report_name="Rapport"):
    ws.set_tab_color(DARK_GREY)
    ws.set_zoom(90)

    df = data["df"]

    col_info = [
        ("date",         "Dato",                    14),
        ("year",         "År",                       8),
        ("month",        "Måned",                    8),
        ("quarter",      "Kvartal",                  9),
        ("brand",        "Varemerke / Sektor",       24),
        ("units",        "Antall",                  10),
        ("net_sales",    "Nettoomsetning (NOK)",     16),
        ("article_no",   "Art.nr.",                  20),
        ("article_desc", "Artikkelbeskrivelse",      32),
    ]

    for ci, (_, _, w) in enumerate(col_info):
        ws.set_column(ci, ci, w)

    ws.set_row(0, 32)
    ws.merge_range(0, 0, 0, len(col_info)-1,
                   f"  {report_name}  |  Data  —  Rensede transaksjonsdata",
                   fmt["title"])

    ws.set_row(1, 20)
    for ci, (_, hdr, _) in enumerate(col_info):
        ws.write(1, ci, hdr, fmt["col_hdr"])

    money_cols = {"net_sales"}
    int_cols   = {"units", "year", "month", "quarter"}

    for ri, (_, row) in enumerate(df.iterrows()):
        is_alt = ri % 2 == 1
        ws.set_row(2 + ri, 15)
        for ci, (col, _, _) in enumerate(col_info):
            val = row.get(col, "")
            if col in money_cols:
                fk = fmt["cell_nok_a"] if is_alt else fmt["cell_nok"]
            elif col in int_cols:
                fk = fmt["cell_int_a"] if is_alt else fmt["cell_int"]
            else:
                fk = fmt["cell_a"] if is_alt else fmt["cell"]
            ws.write(2 + ri, ci, val, fk)

    ws.freeze_panes(2, 0)


# ──────────────────────────────────────────────────────────────────────────────
# Offentlig inngangspunkt
# ──────────────────────────────────────────────────────────────────────────────
def generate_dashboard(data: dict, report_name: str = "Rapport") -> bytes:
    """
    Genererer nettoomsetningsrapporten med 7 regneark.
    Bruker dict fra data_processor.process().
    report_name brukes i alle arkstitler og overskrifter.
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

    _build_dashbord(wb,          ws1, data, fmt, report_name)
    _build_trendanalyse(wb,      ws2, data, fmt, report_name)
    _build_varemerkeanalyse(wb,  ws3, data, fmt, report_name)
    _build_xyz_analyse(wb,       ws4, data, fmt, report_name)
    _build_portfolje(wb,         ws5, data, fmt, report_name)
    _build_topp_artikler(wb,     ws6, data, fmt, report_name)
    _build_data(wb,              ws7, data, fmt, report_name)

    ws1.activate()
    wb.close()
    output.seek(0)
    return output.read()
