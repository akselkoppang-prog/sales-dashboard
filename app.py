"""
app.py  –  Executive Sales Dashboard Generator
Run with:  streamlit run app.py
"""
import io
import time
from datetime import datetime

import pandas as pd
import streamlit as st

from data_processor import process, compute_kpis
from excel_generator import generate_dashboard

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Salgs­dashbord­generator",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    /* Main background */
    .stApp { background-color: #F7F9FC; }

    /* Title block */
    .hero-title {
        font-size: 2.4rem;
        font-weight: 800;
        color: #1A3A5C;
        margin-bottom: 0;
        letter-spacing: -0.5px;
    }
    .hero-subtitle {
        font-size: 1rem;
        color: #595959;
        margin-top: 4px;
        margin-bottom: 28px;
    }

    /* Card wrapper */
    .card {
        background: white;
        border-radius: 12px;
        padding: 28px 32px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.07);
        margin-bottom: 20px;
    }

    /* Section labels */
    .section-label {
        font-size: 0.8rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #2E75B6;
        margin-bottom: 8px;
    }

    /* KPI row */
    .kpi-grid {
        display: flex;
        gap: 16px;
        flex-wrap: wrap;
        margin: 16px 0;
    }
    .kpi-tile {
        flex: 1;
        min-width: 130px;
        background: #EAF3FB;
        border-radius: 10px;
        padding: 16px 12px;
        text-align: center;
        border-left: 4px solid #2E75B6;
    }
    .kpi-label {
        font-size: 0.72rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        color: #595959;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 1.6rem;
        font-weight: 800;
        color: #1A3A5C;
    }
    .kpi-sub {
        font-size: 0.72rem;
        color: #888;
        margin-top: 2px;
    }

    /* Upload zone overrides */
    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #2E75B6 !important;
        border-radius: 10px !important;
        background: #F0F6FC !important;
    }

    /* Generate button */
    div.stButton > button {
        width: 100%;
        padding: 16px;
        font-size: 1.1rem;
        font-weight: 700;
        background: linear-gradient(135deg, #1A3A5C 0%, #2E75B6 100%);
        color: white;
        border: none;
        border-radius: 10px;
        cursor: pointer;
        transition: opacity 0.2s;
    }
    div.stButton > button:hover { opacity: 0.88; }

    /* Download button */
    [data-testid="stDownloadButton"] > button {
        width: 100%;
        padding: 14px;
        font-size: 1.05rem;
        font-weight: 700;
        background: linear-gradient(135deg, #1E8449 0%, #27AE60 100%);
        color: white;
        border: none;
        border-radius: 10px;
    }

    /* Success box */
    .success-box {
        background: #EAFAF1;
        border-left: 5px solid #1E8449;
        border-radius: 8px;
        padding: 16px 20px;
        margin: 16px 0;
        color: #1E8449;
        font-weight: 600;
    }

    /* Error box */
    .error-box {
        background: #FDEDEC;
        border-left: 5px solid #C0392B;
        border-radius: 8px;
        padding: 16px 20px;
        margin: 16px 0;
        color: #C0392B;
        font-weight: 600;
    }

    /* Schema hint */
    .schema-hint {
        font-size: 0.8rem;
        color: #666;
        background: #F2F2F2;
        border-radius: 6px;
        padding: 10px 14px;
        margin-top: 10px;
    }
    .schema-hint code {
        background: #E0E0E0;
        border-radius: 3px;
        padding: 1px 5px;
        font-size: 0.78rem;
    }

    footer { visibility: hidden; }
    #MainMenu { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Hero header
# ---------------------------------------------------------------------------
st.markdown('<div class="hero-title">📊 Salgsrapport­generator</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="hero-subtitle">Last opp salgsdata og generer en profesjonelt formatert '
    'Excel-rapport — kjøres lokalt, ingen data sendes ut av maskinen din.</div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Upload card
# ---------------------------------------------------------------------------
st.markdown('<div class="section-label">1. Last opp data</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    label="Slipp CSV- eller Excel-filen her",
    type=["csv", "xlsx", "xls"],
    label_visibility="collapsed",
)

st.markdown("""
<div class="schema-hint">
    <strong>Forventede kolonner (navn er fleksible):</strong><br>
    <code>Date</code> / <code>Dato</code> &nbsp;·&nbsp;
    <code>Product</code> / <code>Artikkel.Vare modell nr</code> &nbsp;·&nbsp;
    <code>Region</code> / <code>Varusektor</code> &nbsp;·&nbsp;
    <code>Revenue</code> / <code>Omsetning</code> (påkrevd) &nbsp;·&nbsp;
    <code>Quantity</code> / <code>Antall</code><br>
    <em>Manglende valgfrie kolonner fylles inn automatisk. Jærprint-stilfiler gjenkjennes automatisk.</em>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Preview + KPI peek once file is loaded
# ---------------------------------------------------------------------------
df_clean = None
_data = None

if uploaded_file is not None:
    with st.spinner("Leser og renser data…"):
        try:
            _data = process(uploaded_file)
            df_clean = _data["df"]
        except Exception as e:
            st.markdown(f'<div class="error-box">⚠ {e}</div>', unsafe_allow_html=True)

    if df_clean is not None:
        st.markdown("---")
        st.markdown('<div class="section-label">Dataforhåndsvisning</div>', unsafe_allow_html=True)
        preview_cols = ["date", "brand", "net_sales", "units", "article_desc"]
        st.dataframe(
            df_clean[[c for c in preview_cols if c in df_clean.columns]].head(8),
            use_container_width=True,
            hide_index=True,
        )

        kpis = compute_kpis(df_clean)

        def _fmt_money(v):
            if v >= 1_000_000:
                return f"NOK {v/1_000_000:.1f}M"
            if v >= 1_000:
                return f"NOK {v/1_000:.1f}K"
            return f"NOK {v:,.0f}"

        mom_html = ""
        if kpis["mom_growth"] is not None:
            sign = "+" if kpis["mom_growth"] >= 0 else ""
            color = "#1E8449" if kpis["mom_growth"] >= 0 else "#C0392B"
            mom_html = f"""
            <div class="kpi-tile" style="border-left-color:{color}">
                <div class="kpi-label">MoM-vekst</div>
                <div class="kpi-value" style="color:{color}">{sign}{kpis['mom_growth']:.1f}%</div>
            </div>"""

        st.markdown(f"""
        <div class="kpi-grid">
            <div class="kpi-tile">
                <div class="kpi-label">Total omsetning</div>
                <div class="kpi-value">{_fmt_money(kpis['total_revenue'])}</div>
            </div>
            <div class="kpi-tile">
                <div class="kpi-label">Snitt ordrebeløp</div>
                <div class="kpi-value">{_fmt_money(kpis['avg_order_value'])}</div>
            </div>
            <div class="kpi-tile">
                <div class="kpi-label">Transaksjoner</div>
                <div class="kpi-value">{kpis['num_transactions']:,}</div>
            </div>
            <div class="kpi-tile">
                <div class="kpi-label">Totalt antall solgt</div>
                <div class="kpi-value">{kpis['total_quantity']:,}</div>
            </div>
            {mom_html}
        </div>
        """, unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Generate button
# ---------------------------------------------------------------------------
st.markdown("---")
st.markdown('<div class="section-label">2. Generer rapport</div>', unsafe_allow_html=True)

report_name = st.text_input(
    "Rapport / selskapsnavn",
    value="Jærprint",
    placeholder="f.eks. Jærprint, Acme AS, ...",
    help="Dette navnet vises i alle arkstitler og overskrifter i den genererte Excel-filen.",
).strip() or "Rapport"

if df_clean is None:
    st.button("Generer lederrapport", disabled=True)
else:
    if st.button("Generer lederrapport"):
        with st.spinner("Bygger Excel-rapport…"):
            try:
                t0 = time.time()
                excel_bytes = generate_dashboard(_data, report_name=report_name)
                elapsed = time.time() - t0

                safe_name = report_name.replace(" ", "_").replace("/", "-")
                filename = f"{safe_name}_rapport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

                st.markdown(
                    f'<div class="success-box">✓ Rapport generert på {elapsed:.1f}s — '
                    f'{len(df_clean):,} rader behandlet fordelt på 7 ark.</div>',
                    unsafe_allow_html=True,
                )

                st.download_button(
                    label="⬇  Last ned lederrapport (.xlsx)",
                    data=excel_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.markdown("""
                **Rapporten inneholder:**
                - **Ark 1 – Dashbord:** 7 KPI-fliser, ledersammendrag, årssammendrag + strategiske innsikter (HHI, ABC, Pareto, Gini, sesong)
                - **Ark 2 – Trendanalyse:** Månedlig og kvartalsvis omsetning, ÅoÅ-vekst %, sesongindeks + linjediagram
                - **Ark 3 – Varemerkeanalyse:** Omsetning per varemerke med ÅoÅ, andel og ABC-klassifisering, Pareto 80/20 + stolpediagram
                - **Ark 4 – XYZ-analyse:** Etterspørselsvariabilitet (CV), ABC–XYZ-matrise per varemerke
                - **Ark 5 – Portefølje:** BCG-inspirert vekst/andel-matrise (Stjerne, Melkeku, Spørsmålstegn, Hund)
                - **Ark 6 – Topp-artikler:** Bestselgende artikler per antall for topp 5 varemerker
                - **Ark 7 – Data:** Rensede transaksjonsdata med frosne overskrifter
                """)

            except Exception as e:
                st.markdown(f'<div class="error-box">⚠ Feil ved generering av rapport: {e}</div>', unsafe_allow_html=True)
                st.exception(e)

# ---------------------------------------------------------------------------
# Footer hint
# ---------------------------------------------------------------------------
st.markdown("---")
st.markdown(
    "<small style='color:#aaa;'>Kjøres 100% lokalt — ingen data sendes til noen server.</small>",
    unsafe_allow_html=True,
)
