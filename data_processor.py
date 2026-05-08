"""
data_processor.py
Loads a Jærprint-style sales file and produces all aggregations needed
for the four-sheet Executive Dashboard.
"""
import re
import math
import pandas as pd
import numpy as np
from datetime import datetime

# ── Column mapping (Norwegian → internal names) ─────────────────────────────
_COL_MAP = {
    "dato":                    "date",
    "varusektor":              "brand",
    "antall":                  "units",
    "omsetning":               "net_sales",
    "artikkel.vare_modell_nr": "article_no",
    "artikel":                 "article_desc",
    # English fallbacks
    "date":        "date",
    "brand":       "brand",
    "units":       "units",
    "revenue":     "net_sales",
    "net_sales":   "net_sales",
    "article_no":  "article_no",
    "article_desc":"article_desc",
    "product":     "article_no",
    "quantity":    "units",
    "region":      "brand",
}

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]


_SIZE_RE = re.compile(
    r'\s+(?:XXS|XXL|XXXL|3XL|4XL|5XL|2XL|XS|XL|S|M|L|\d{1,3}XL|\d{2,3})\s*$',
    re.IGNORECASE,
)


def _base_article_desc(descs) -> str:
    """
    Given a list of description strings for the same base article number
    (e.g. ['029030-55-9 Basic-T Royal Blue 3XL', '029030-60-9 Basic-T Navy S']),
    return the base product name stripped of article-no prefix, colour, and size.

    Strategy:
      1. Strip the leading article-number token (first whitespace-delimited word).
      2. Strip trailing standard size code.
      3. Find the longest common word-prefix across all variants — this naturally
         drops colour words that differ across variants.
      4. Fall back to the first size-stripped description if prefix is empty.
    """
    clean = []
    for d in descs:
        s = str(d).strip()
        if not s or s.lower() == "nan":
            continue
        # Strip leading article-number token ("029030-55-9 ...")
        s = re.sub(r'^\S+\s+', '', s).strip()
        # Strip trailing size code
        s = _SIZE_RE.sub('', s).strip()
        if s:
            clean.append(s)

    if not clean:
        return ""
    if len(clean) == 1:
        return clean[0]

    # Find common word-prefix
    words = [s.split() for s in clean]
    common = []
    for i, word in enumerate(words[0]):
        if all(i < len(w) and w[i] == word for w in words[1:]):
            common.append(word)
        else:
            break
    return " ".join(common) if common else clean[0]


def _base_article_no(code: str) -> str:
    """
    Strip trailing colour-code and size-code suffixes from an article number so
    that different colour/size variants of the same product are grouped together.

    Rule: if the last TWO hyphen-separated segments are both purely numeric,
    they are treated as colour-code and size-code and removed.
      '029030-55-9'  → '029030'   (base-colour-size)
      '029030-55-10' → '029030'
    Codes that don't match the pattern are returned unchanged.
      'ART-001'      → 'ART-001'
      'ABC-XL'       → 'ABC-XL'
    """
    parts = code.split("-")
    if len(parts) >= 3 and parts[-1].isdigit() and parts[-2].isdigit():
        return "-".join(parts[:-2])
    return code


def _normalise(df: pd.DataFrame) -> pd.DataFrame:
    norm = {c: c.lower().replace(" ", "_") for c in df.columns}
    df = df.rename(columns=norm)
    rename = {}
    for col in df.columns:
        if col in _COL_MAP:
            rename[col] = _COL_MAP[col]
    return df.rename(columns=rename)


def process(uploaded_file):
    """
    Load, clean, and aggregate the raw sales file.
    Returns a dict with all data structures needed by the Excel generator.
    """
    name = getattr(uploaded_file, "name", "").lower()
    if name.endswith(".csv"):
        raw = pd.read_csv(uploaded_file)
    else:
        raw = pd.read_excel(uploaded_file)

    df = _normalise(raw)

    # Drop summary / totals rows (no parseable date)
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df[df["date"].notna()].copy()

    # Ensure required columns exist
    if "brand"    not in df.columns: df["brand"]    = "Unknown"
    if "units"    not in df.columns: df["units"]    = 1
    if "net_sales"not in df.columns:
        raise ValueError("Cannot find a Net Sales / Revenue column.")
    if "article_no"   not in df.columns: df["article_no"]   = ""
    if "article_desc" not in df.columns: df["article_desc"] = ""

    # Clean numeric columns
    for col in ["net_sales", "units"]:
        df[col] = (pd.to_numeric(
            df[col].astype(str).str.replace(r"[^\d.\-]", "", regex=True),
            errors="coerce"
        ).fillna(0))

    df["units"] = df["units"].astype(int)
    df["brand"]       = df["brand"].astype(str).str.strip()
    df["article_no"]  = df["article_no"].astype(str).str.strip()
    df["article_desc"]= df["article_desc"].astype(str).str.strip()

    df["year"]    = df["date"].dt.year.astype(int)
    df["month"]   = df["date"].dt.month.astype(int)
    df["quarter"] = df["date"].dt.quarter.astype(int)

    df = df.sort_values("date").reset_index(drop=True)

    # ── Reporting dimensions ────────────────────────────────────────────────
    all_years  = sorted(df["year"].unique())
    max_year   = max(all_years)
    max_month  = int(df[df["year"] == max_year]["month"].max())
    ytd_months = list(range(1, max_month + 1))
    ytd_label  = f"Jan–{MONTHS[max_month - 1]}"

    full_years = [y for y in all_years if y < max_year]  # complete calendar years

    # ── Annual totals ───────────────────────────────────────────────────────
    annual = df.groupby("year")["net_sales"].sum()

    # ── KPIs ───────────────────────────────────────────────────────────────
    fy_vals  = {y: float(annual.get(y, 0)) for y in all_years}
    ytd_vals = {
        y: float(df[(df["year"] == y) & (df["month"].isin(ytd_months))]["net_sales"].sum())
        for y in all_years
    }

    # CAGR over full years (skip partial year)
    cagr = None
    if len(full_years) >= 2:
        y0, y1 = full_years[0], full_years[-1]
        n = y1 - y0
        if fy_vals[y0] > 0:
            cagr = (fy_vals[y1] / fy_vals[y0]) ** (1 / n) - 1

    # Brand all-time totals
    brand_all = df.groupby("brand")["net_sales"].sum().sort_values(ascending=False)
    grand_total = float(brand_all.sum())
    top1_brand  = brand_all.index[0] if len(brand_all) else "N/A"
    top1_share  = float(brand_all.iloc[0] / grand_total) if grand_total else 0
    top3_share  = float(brand_all.iloc[:3].sum() / grand_total) if grand_total else 0

    # ── Annual summary table ────────────────────────────────────────────────
    prev_ytd = None
    ann_rows = []
    for y in all_years:
        ns   = fy_vals[y]
        prev = fy_vals.get(y - 1, None)
        yoy  = (ns - prev) / prev if (prev and prev != 0) else None
        ytd  = ytd_vals[y]
        prev_ytd_v = ytd_vals.get(y - 1, None)
        ytd_yoy = (ytd - prev_ytd_v) / prev_ytd_v if (prev_ytd_v and prev_ytd_v != 0) else None
        # Best quarter
        qtr = df[df["year"] == y].groupby("quarter")["net_sales"].sum()
        best_q = f"Q{int(qtr.idxmax())}" if len(qtr) else "—"
        ann_rows.append({
            "year": str(y), "net_sales": ns,
            "yoy": yoy, "ytd": ytd, "ytd_yoy": ytd_yoy, "best_q": best_q,
        })
    annual_summary = pd.DataFrame(ann_rows)

    # ── Monthly pivot ───────────────────────────────────────────────────────
    mpivot = (df.groupby(["year", "month"])["net_sales"].sum()
              .unstack(fill_value=0)
              .reindex(columns=range(1, 13), fill_value=0))
    mpivot.index = mpivot.index.astype(str)
    mpivot.columns = MONTHS
    mpivot["TOTAL"] = mpivot.sum(axis=1)

    # Add TOTAL row
    tot_row = pd.DataFrame([mpivot.sum()], index=["TOTAL"])
    monthly_pivot = pd.concat([mpivot, tot_row])

    # ── Monthly YoY growth ──────────────────────────────────────────────────
    yoy_rows = []
    for i in range(1, len(all_years)):
        y0, y1 = all_years[i - 1], all_years[i]
        row = {"label": f"{y1} vs {y0}"}
        for mi, m in enumerate(MONTHS):
            v0 = float(mpivot.loc[str(y0), m]) if str(y0) in mpivot.index else 0
            v1 = float(mpivot.loc[str(y1), m]) if str(y1) in mpivot.index else 0
            row[m] = (v1 - v0) / v0 if v0 != 0 else (None if v1 == 0 else None)
        t0 = float(mpivot.loc[str(y0), "TOTAL"]) if str(y0) in mpivot.index else 0
        t1 = float(mpivot.loc[str(y1), "TOTAL"]) if str(y1) in mpivot.index else 0
        row["Full Year"] = (t1 - t0) / t0 if t0 != 0 else None
        yoy_rows.append(row)
    monthly_yoy = pd.DataFrame(yoy_rows)

    # ── Quarterly pivot ─────────────────────────────────────────────────────
    qpivot = (df.groupby(["year", "quarter"])["net_sales"].sum()
              .unstack(fill_value=0)
              .reindex(columns=[1, 2, 3, 4], fill_value=0))
    qpivot.index = qpivot.index.astype(str)
    qpivot.columns = ["Q1", "Q2", "Q3", "Q4"]
    qpivot["Total"] = qpivot.sum(axis=1)
    qpivot["Q2 Share"] = qpivot["Q2"] / qpivot["Total"].replace(0, np.nan)
    qpivot = qpivot.reset_index().rename(columns={"year": "Year"})

    # ── Seasonality index ────────────────────────────────────────────────────
    # Based on full years only (exclude partial current year)
    df_full = df[df["year"].isin(full_years)]
    if len(df_full) > 0:
        monthly_avg_per_year = (
            df_full.groupby(["year", "month"])["net_sales"].sum()
            .groupby("month").mean()
        )
        overall_avg = monthly_avg_per_year.mean()
        seas_index = (monthly_avg_per_year / overall_avg).round(2)
    else:
        seas_index = pd.Series([1.0] * 12, index=range(1, 13))
    seas_dict = {MONTHS[m - 1]: float(seas_index.get(m, 0)) for m in range(1, 13)}

    # ── Brand performance ────────────────────────────────────────────────────
    brand_years = df.groupby(["brand", "year"])["net_sales"].sum().unstack(fill_value=0)
    for y in all_years:
        if y not in brand_years.columns:
            brand_years[y] = 0.0
    brand_years = brand_years[all_years]

    last_full = full_years[-1] if full_years else max_year
    brand_years = brand_years.sort_values(last_full, ascending=False)

    brand_ytd = df[df["month"].isin(ytd_months)].groupby(["brand", "year"])["net_sales"].sum().unstack(fill_value=0)

    brand_rows = []
    total_last_full = float(brand_years[last_full].sum())
    for brand in brand_years.index:
        row = {"brand": brand}
        for y in all_years:
            row[f"FY{y}"] = float(brand_years.loc[brand, y]) if y in brand_years.columns else 0.0
        # YTD for max year
        row["YTD"] = float(brand_ytd.loc[brand, max_year]) if (brand in brand_ytd.index and max_year in brand_ytd.columns) else 0.0
        # YoY
        for i in range(1, len(all_years)):
            y0, y1 = all_years[i - 1], all_years[i]
            v0 = row.get(f"FY{y0}", 0)
            v1 = row.get(f"FY{y1}", 0)
            if v0 == 0:
                row[f"YoY_{y1}v{y0}"] = None  # new brand
            elif v1 == 0:
                row[f"YoY_{y1}v{y0}"] = -1.0
            else:
                row[f"YoY_{y1}v{y0}"] = (v1 - v0) / v0
        # Share of last full year
        v_last = row.get(f"FY{last_full}", 0)
        row["share"] = (v_last / total_last_full) if total_last_full > 0 else 0
        brand_rows.append(row)

    brand_perf = pd.DataFrame(brand_rows)

    # Add TOTAL row
    total_row = {"brand": "TOTAL"}
    for y in all_years:
        total_row[f"FY{y}"] = float(brand_years[y].sum())
    total_row["YTD"] = float(brand_ytd[max_year].sum()) if max_year in brand_ytd.columns else 0
    for i in range(1, len(all_years)):
        total_row[f"YoY_{all_years[i]}v{all_years[i-1]}"] = None
    total_row["share"] = None
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", FutureWarning)
        brand_perf = pd.concat(
            [brand_perf, pd.DataFrame([{c: total_row.get(c, np.nan) for c in brand_perf.columns}])],
            ignore_index=True,
        )

    # ── Pareto analysis ──────────────────────────────────────────────────────
    pareto_df = brand_all.reset_index()
    pareto_df.columns = ["brand", "total_revenue"]
    pareto_df["pct"] = pareto_df["total_revenue"] / grand_total
    pareto_df["cumulative"] = pareto_df["pct"].cumsum()
    pareto_df["rank"] = range(1, len(pareto_df) + 1)
    pareto_df["threshold_80"] = pareto_df["cumulative"].apply(
        lambda x: "80%" if x >= 0.80 else ""
    )

    # ── ABC brand classification (by last full year revenue) ─────────────────
    # A = top brands covering 0–70% cumulative revenue
    # B = 70–90%   C = 90–100%
    # (Standard ABC / selective inventory control — Dickie 1951)
    abc_brands = {}
    if total_last_full > 0:
        brand_lfy = (df[df["year"] == last_full]
                     .groupby("brand")["net_sales"].sum()
                     .sort_values(ascending=False))
        cum_share = brand_lfy.cumsum() / brand_lfy.sum()
        prev_cum = 0.0
        for brand_name, cum_val in cum_share.items():
            if prev_cum < 0.70:
                abc_brands[brand_name] = "A"
            elif prev_cum < 0.90:
                abc_brands[brand_name] = "B"
            else:
                abc_brands[brand_name] = "C"
            prev_cum = float(cum_val)

    # ── HHI — Herfindahl-Hirschman Index ────────────────────────────────────
    # Measures portfolio concentration risk (0 = perfect diversity, 1 = monopoly)
    # Thresholds: >0.25 highly concentrated, 0.15–0.25 moderate, <0.15 diverse
    brand_shares_all = brand_all / grand_total if grand_total > 0 else brand_all * 0
    hhi = float((brand_shares_all ** 2).sum())

    # ── Seasonality: peak and trough months ──────────────────────────────────
    peak_month   = max(seas_dict, key=seas_dict.get) if seas_dict else "N/A"
    trough_month = min(seas_dict, key=seas_dict.get) if seas_dict else "N/A"

    # ── XYZ demand variability analysis ──────────────────────────────────────
    # Coefficient of Variation (CV) = std / mean of monthly brand sales
    # X: CV < 0.50 (stabil)  Y: 0.50–1.00 (variabel)  Z: > 1.00 (erratisk)
    # Ref: Silver, Pyke & Peterson (1998); Scholz-Reiter et al. (2012)
    monthly_brand = df.groupby(["brand", "year", "month"])["net_sales"].sum().reset_index()
    xyz_brands = {}
    xyz_rows = []
    for b_name in brand_all.index:
        bs = monthly_brand[monthly_brand["brand"] == b_name]["net_sales"]
        n_obs = len(bs)
        if n_obs >= 3:
            mean_s = float(bs.mean())
            std_s  = float(bs.std(ddof=1))
            cv     = std_s / mean_s if mean_s > 0 else 0.0
        else:
            mean_s = float(brand_all.get(b_name, 0)) / max(1, n_obs)
            std_s  = 0.0
            cv     = 0.0
        xyz = "X" if cv < 0.5 else ("Y" if cv < 1.0 else "Z")
        xyz_brands[b_name] = xyz
        xyz_rows.append({
            "brand":         b_name,
            "mean_monthly":  mean_s,
            "std_monthly":   std_s,
            "cv":            cv,
            "n_months":      n_obs,
            "xyz":           xyz,
            "abc":           abc_brands.get(b_name, "—"),
            "combined":      abc_brands.get(b_name, "—") + xyz if abc_brands.get(b_name) else "—",
        })
    xyz_df = pd.DataFrame(xyz_rows)

    # ABC–XYZ matrix count (3×3 grid)
    abc_xyz_matrix = {
        abc: {xyz: 0 for xyz in ["X", "Y", "Z"]}
        for abc in ["A", "B", "C"]
    }
    for row in xyz_rows:
        a = row["abc"]
        x = row["xyz"]
        if a in abc_xyz_matrix and x in abc_xyz_matrix[a]:
            abc_xyz_matrix[a][x] += 1

    # ── Portfolio matrix (BCG-inspirert) ─────────────────────────────────────
    # Kvadranter: Stjerne (høy vekst + høy andel), Melkeku, Spørsmålstegn, Hund
    # Ref: Henderson (1970); tilpasset intern merkeporteføljeanalyse
    portfolio_df = pd.DataFrame()
    portfolio_avg_growth = None
    portfolio_avg_share  = None
    portfolio_ref_years  = None
    if len(full_years) >= 2:
        lfy_p  = full_years[-1]
        prev_p = full_years[-2]
        avg_share_thr = 1.0 / max(1, len(brand_all))
        growth_map = {}
        for b_name in brand_all.index:
            v0 = float(df[(df["brand"] == b_name) & (df["year"] == prev_p)]["net_sales"].sum())
            v1 = float(df[(df["brand"] == b_name) & (df["year"] == lfy_p)]["net_sales"].sum())
            growth_map[b_name] = (v1 - v0) / v0 if v0 > 0 else None
        valid_g = [g for g in growth_map.values() if g is not None]
        avg_g   = float(np.mean(valid_g)) if valid_g else 0.0
        p_rows  = []
        for b_name in brand_all.index:
            share  = float(brand_shares_all.get(b_name, 0))
            growth = growth_map.get(b_name)
            if growth is None:
                cat = "Ny"
            elif growth >= avg_g and share >= avg_share_thr:
                cat = "Stjerne"
            elif growth < avg_g and share >= avg_share_thr:
                cat = "Melkeku"
            elif growth >= avg_g and share < avg_share_thr:
                cat = "Spørsmålstegn"
            else:
                cat = "Hund"
            p_rows.append({
                "brand":      b_name,
                "share_pct":  share,
                "growth_pct": growth,
                "category":   cat,
                "abc":        abc_brands.get(b_name, "—"),
            })
        portfolio_df         = pd.DataFrame(p_rows)
        portfolio_avg_growth = avg_g
        portfolio_avg_share  = avg_share_thr
        portfolio_ref_years  = (prev_p, lfy_p)

    # ── Gini coefficient (omsetningsfordeling) ────────────────────────────────
    # 0 = perfekt likhet, 1 = all omsetning hos ett varemerke
    # Ref: Gini (1912)
    _sv = np.sort(brand_all.values)
    _n  = len(_sv)
    if _n > 1 and _sv.sum() > 0:
        _idx = np.arange(1, _n + 1)
        gini = float((2 * np.sum(_idx * _sv)) / (_n * _sv.sum()) - (_n + 1) / _n)
    else:
        gini = 0.0

    # ── Top articles per top-5 brand (by units sold, base article code) ─────
    # Colour/size variants are consolidated: '029030-55-9' and '029030-55-10'
    # both map to base code '029030' before grouping.
    top5_brands = brand_all.nlargest(5).index.tolist()
    top_items_per_brand = {}
    for b in top5_brands:
        bdf = df[df["brand"] == b].copy()
        bdf["base_no"] = bdf["article_no"].apply(_base_article_no)
        # Consolidated description: strip article-no prefix, colour, and size words
        desc_map = (bdf.groupby("base_no")["article_desc"]
                    .apply(lambda x: _base_article_desc(x.tolist()))
                    .to_dict())
        items = (bdf.groupby("base_no")
                 .agg(units=("units", "sum"), net_sales=("net_sales", "sum"))
                 .reset_index()
                 .rename(columns={"base_no": "article_no"})
                 .sort_values("units", ascending=False)
                 .head(8)
                 .reset_index(drop=True))
        items["article_desc"] = items["article_no"].map(desc_map).fillna("")
        top_items_per_brand[b] = items

    # ── Clean data sheet ─────────────────────────────────────────────────────
    df_export = df[["date", "year", "month", "quarter", "brand", "units", "net_sales", "article_no", "article_desc"]].copy()
    df_export["date"] = df_export["date"].dt.strftime("%Y-%m-%d")

    return {
        "df":                  df_export,
        "all_years":           all_years,
        "max_year":            max_year,
        "max_month":           max_month,
        "ytd_label":           ytd_label,
        "full_years":          full_years,
        "fy_vals":             fy_vals,
        "ytd_vals":            ytd_vals,
        "cagr":                cagr,
        "top1_brand":          top1_brand,
        "top1_share":          top1_share,
        "top3_share":          top3_share,
        "grand_total":         grand_total,
        "annual_summary":      annual_summary,
        "monthly_pivot":       monthly_pivot,
        "monthly_yoy":         monthly_yoy,
        "quarterly_pivot":     qpivot,
        "seasonality":         seas_dict,
        "brand_perf":          brand_perf,
        "pareto":              pareto_df,
        "last_full_year":      last_full,
        "abc_brands":          abc_brands,
        "hhi":                 hhi,
        "peak_month":          peak_month,
        "trough_month":        trough_month,
        "top_items_per_brand":  top_items_per_brand,
        "top5_brands":          top5_brands,
        "xyz_brands":           xyz_brands,
        "xyz_df":               xyz_df,
        "abc_xyz_matrix":       abc_xyz_matrix,
        "portfolio_df":         portfolio_df,
        "portfolio_avg_growth": portfolio_avg_growth,
        "portfolio_avg_share":  portfolio_avg_share,
        "portfolio_ref_years":  portfolio_ref_years,
        "gini":                 gini,
    }


def compute_kpis(df_clean: pd.DataFrame) -> dict:
    """Lightweight KPI dict for the Streamlit preview (operates on df_export)."""
    total = float(df_clean["net_sales"].sum())
    units = int(df_clean["units"].sum())
    avg   = total / len(df_clean) if len(df_clean) else 0
    brands = df_clean["brand"].nunique()
    return {
        "total_revenue": total, "total_quantity": units,
        "avg_order_value": avg, "num_transactions": len(df_clean),
        "unique_products": df_clean["article_no"].nunique(),
        "unique_reps": brands, "has_salesrep": False, "mom_growth": None,
    }


def load_and_clean(uploaded_file) -> pd.DataFrame:
    """Compatibility shim — returns the cleaned transaction df."""
    data = process(uploaded_file)
    df = data["df"].copy()
    df = df.rename(columns={
        "date": "date", "brand": "region", "units": "quantity",
        "net_sales": "revenue", "article_no": "product",
        "article_desc": "product_label",
    })
    df["salesrep"] = "N/A"
    df["date"] = pd.to_datetime(df["date"])
    df["year"] = df["date"].dt.year
    df["month"] = df["date"].dt.to_period("M")
    df["quarter"] = df["date"].dt.to_period("Q")
    df["month_label"] = df["date"].dt.strftime("%b %Y")
    df["quarter_label"] = df["quarter"].astype(str)
    return df
