"""
app.py — Financial Forecasting Web App
Upload financial statement PDFs → Claude extracts data → Download Excel model

Run locally:   streamlit run app.py
Deploy:        Push to GitHub, connect repo on share.streamlit.io
"""

import io
import os
import zipfile
import streamlit as st
from extractor import extract_financials, normalize_extracted_data
from excel_builder import build_period_file, build_forecast_model

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Financial Forecasting Tool",
    page_icon="📊",
    layout="wide",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { max-width: 900px; }
    .stProgress > div > div > div > div { background-color: #1F4E79; }
    .success-box {
        background-color: #E2EFDA;
        border-left: 4px solid #008000;
        padding: 12px 16px;
        border-radius: 4px;
        margin: 8px 0;
    }
    .error-box {
        background-color: #FFE7E7;
        border-left: 4px solid #CC0000;
        padding: 12px 16px;
        border-radius: 4px;
        margin: 8px 0;
    }
    .metric-card {
        background-color: #F5F5F5;
        padding: 12px;
        border-radius: 6px;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)


# ── Helpers ────────────────────────────────────────────────────────────────────
def fmt_currency(val, currency="$"):
    if val is None:
        return "N/A"
    if abs(val) >= 1_000_000:
        return f"{currency}{val/1_000_000:.1f}M"
    if abs(val) >= 1_000:
        return f"{currency}{val/1_000:.0f}K"
    return f"{currency}{val:,.0f}"

def fmt_pct(val):
    if val is None:
        return "N/A"
    return f"{val:.1%}"

def create_zip(period_files: dict, model_bytes: bytes) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, data in period_files.items():
            zf.writestr(f"processed/{filename}", data)
        zf.writestr("model/forecast_model.xlsx", model_bytes)
    return buf.getvalue()


# ── Header ─────────────────────────────────────────────────────────────────────
st.title("📊 Financial Forecasting Tool")
st.markdown("Upload financial statement PDFs → Extract data with Claude AI → Download forecast model in Excel")
st.divider()

# ── Sidebar: API Key ───────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        value=os.environ.get("ANTHROPIC_API_KEY", ""),
        help="Get your key at console.anthropic.com. You can also set ANTHROPIC_API_KEY as an env variable.",
        placeholder="sk-ant-...",
    )

    model = st.selectbox(
        "Claude Model",
        ["claude-sonnet-4-6", "claude-opus-4-6", "claude-haiku-4-5-20251001"],
        index=0,
        help="Sonnet is the best balance of speed and accuracy for extraction.",
    )

    st.divider()
    st.markdown("**How it works:**")
    st.markdown("1. Upload 1–10 PDF financial statements")
    st.markdown("2. Claude reads each PDF and extracts all financial data")
    st.markdown("3. Download per-period Excel files + a forecast model")
    st.divider()
    st.markdown("**Tips:**")
    st.markdown("- Works with annual reports, quarterly statements, any company")
    st.markdown("- Handles text-based PDFs best; scanned PDFs may have lower accuracy")
    st.markdown("- Upload multiple years for a richer forecast model")


# ── Main: File Upload ──────────────────────────────────────────────────────────
col1, col2 = st.columns([2, 1])

with col1:
    uploaded_files = st.file_uploader(
        "Upload Financial Statement PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        help="Upload one PDF per reporting period (e.g. one per year, or one per quarter)",
    )

with col2:
    if uploaded_files:
        st.metric("Files uploaded", len(uploaded_files))
    else:
        st.info("Upload PDFs to get started")

if not uploaded_files:
    st.stop()

if not api_key:
    st.warning("⚠️ Please enter your Anthropic API key in the sidebar to proceed.")
    st.stop()


# ── Process Button ─────────────────────────────────────────────────────────────
st.divider()
run_col, _ = st.columns([1, 3])
with run_col:
    run = st.button("🚀 Extract & Build Model", type="primary", use_container_width=True)

if not run:
    st.stop()


# ── Extraction ─────────────────────────────────────────────────────────────────
st.divider()
st.subheader("📥 Extracting Financial Data")

results = []
period_files = {}
errors = []

progress_bar = st.progress(0)
status_text  = st.empty()

for idx, uploaded_file in enumerate(uploaded_files):
    filename = uploaded_file.name
    status_text.text(f"Processing {filename}... ({idx+1}/{len(uploaded_files)})")

    pdf_bytes = uploaded_file.read()

    with st.spinner(f"Claude is reading {filename}..."):
        raw = extract_financials(pdf_bytes, api_key, model)
        raw["_source_filename"] = filename
        data = normalize_extracted_data(raw)

    if "_extraction_error" in data:
        errors.append((filename, data["_extraction_error"]))
        st.markdown(f"""<div class="error-box">❌ <b>{filename}</b>: {data['_extraction_error']}</div>""",
                    unsafe_allow_html=True)
    else:
        results.append(data)
        company  = data.get("company_name", "Unknown")
        period   = data.get("period", "?")
        revenue  = data.get("revenue")
        net_inc  = data.get("net_income")
        currency = data.get("currency", "$")

        st.markdown(f"""<div class="success-box">
            ✅ <b>{filename}</b><br>
            Company: <b>{company}</b> &nbsp;|&nbsp; Period: <b>{period}</b> &nbsp;|&nbsp;
            Revenue: <b>{fmt_currency(revenue, currency)}</b> &nbsp;|&nbsp;
            Net Income: <b>{fmt_currency(net_inc, currency)}</b>
        </div>""", unsafe_allow_html=True)

        # Build per-period Excel
        excel_bytes = build_period_file(data)
        period_files[f"{period}_financials.xlsx"] = excel_bytes

    progress_bar.progress((idx + 1) / len(uploaded_files))

status_text.empty()
progress_bar.progress(1.0)


# ── Extraction Summary ─────────────────────────────────────────────────────────
if errors and not results:
    st.error("All files failed to process. Please check your API key and try again.")
    st.stop()

if results:
    st.divider()
    st.subheader("📊 Extracted Data Preview")

    # Summary metrics across all periods
    cols = st.columns(len(results))
    for i, d in enumerate(results):
        with cols[i]:
            rev = d.get("revenue")
            ni  = d.get("net_income")
            gm  = d.get("gross_margin_pct")
            cur = d.get("currency", "$")
            st.markdown(f"**{d.get('period', '?')}**")
            st.metric("Revenue", fmt_currency(rev, cur))
            st.metric("Net Income", fmt_currency(ni, cur))
            st.metric("Gross Margin", fmt_pct(gm))

    # ── Build Forecast Model ─────────────────────────────────────────────────
    st.divider()
    st.subheader("🔮 Building Forecast Model")

    with st.spinner("Building forecast model with Historical_Model, Drivers, Forecast, Scenarios..."):
        model_bytes = build_forecast_model(results)

    st.success(f"✅ Forecast model built from {len(results)} period(s).")

    # ── Downloads ────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("⬇️ Download Outputs")

    dl_col1, dl_col2, dl_col3 = st.columns(3)

    # Per-period files
    with dl_col1:
        st.markdown("**Per-period Excel files**")
        for filename, data_bytes in period_files.items():
            st.download_button(
                label=f"📄 {filename}",
                data=data_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # Forecast model
    with dl_col2:
        st.markdown("**Forecast model**")
        st.download_button(
            label="📊 forecast_model.xlsx",
            data=model_bytes,
            file_name="forecast_model.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # Full ZIP
    with dl_col3:
        st.markdown("**Everything (ZIP)**")
        zip_bytes = create_zip(period_files, model_bytes)
        st.download_button(
            label="🗜️ Download all files",
            data=zip_bytes,
            file_name="financial_forecast.zip",
            mime="application/zip",
            use_container_width=True,
        )

    # ── Extracted data details ───────────────────────────────────────────────
    with st.expander("🔍 View extracted data details"):
        for d in results:
            st.markdown(f"#### {d.get('company_name', '')} — {d.get('period', '')}")
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.markdown("**Income Statement**")
                st.json({k: v for k, v in d.items()
                         if k in ["revenue", "cogs", "gross_profit", "total_opex",
                                  "operating_income", "net_income", "gross_margin_pct"]})
            with col_b:
                st.markdown("**Cash Flow**")
                st.json({k: v for k, v in d.items()
                         if k in ["ocf", "capex", "investing_cf", "financing_cf", "amortization"]})
            with col_c:
                st.markdown("**Balance Sheet**")
                st.json({k: v for k, v in d.items()
                         if k in ["accounts_receivable", "inventories", "capital_assets",
                                  "total_assets", "accounts_payable", "total_debt", "total_equity"]})
