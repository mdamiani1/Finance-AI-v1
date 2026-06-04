"""
app.py  Financial Forecasting Tool  Clean minimal UI
"""
import io, os, zipfile
import streamlit as st
import plotly.graph_objects as go
from extractor import extract_financials, normalize_extracted_data
from excel_builder import build_period_file, build_forecast_model

st.set_page_config(page_title="Finance AI", page_icon="📊", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
#MainMenu,footer,header{visibility:hidden;}
.block-container{padding-top:2rem;padding-bottom:2rem;max-width:1100px;}
.hero{background:linear-gradient(135deg,#0f172a 0%,#1e293b 100%);border-radius:16px;padding:48px 40px;margin-bottom:28px;color:white;}
.hero-badge{display:inline-block;background:rgba(99,102,241,0.2);border:1px solid rgba(99,102,241,0.4);color:#a5b4fc;padding:4px 12px;border-radius:20px;font-size:12px;font-weight:500;letter-spacing:0.5px;margin-bottom:16px;}
.hero h1{font-size:36px;font-weight:700;margin:0 0 12px 0;color:white;letter-spacing:-0.5px;}
.hero p{font-size:16px;color:#94a3b8;margin:0;max-width:560px;line-height:1.6;}
.steps-row{display:flex;gap:12px;margin-bottom:28px;}
.step-card{flex:1;background:white;border:1px solid #e2e8f0;border-radius:12px;padding:20px;}
.step-num{width:28px;height:28px;background:#f1f5f9;border-radius:8px;display:inline-flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;color:#475569;margin-bottom:10px;}
.step-num.active{background:#6366f1;color:white;}
.step-title{font-size:13px;font-weight:600;color:#1e293b;margin-bottom:4px;}
.step-desc{font-size:12px;color:#94a3b8;line-height:1.4;}
.card{background:white;border:1px solid #e2e8f0;border-radius:12px;padding:24px;margin-bottom:16px;}
.card h3{font-size:15px;font-weight:600;color:#1e293b;margin:0 0 4px 0;}
.card p{font-size:13px;color:#64748b;margin:0 0 16px 0;}
.ok{background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:12px 16px;margin:6px 0;font-size:13px;color:#166534;}
.err{background:#fef2f2;border:1px solid #fecaca;border-radius:8px;padding:12px 16px;margin:6px 0;font-size:13px;color:#991b1b;}
.metric-inner{background:white;border:1px solid #e2e8f0;border-radius:12px;padding:16px 20px;margin-top:10px;}
.metric-per{font-size:11px;font-weight:600;color:#94a3b8;letter-spacing:0.5px;text-transform:uppercase;margin-bottom:12px;}
.mrow{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:8px;}
.mlabel{font-size:12px;color:#64748b;}
.mval{font-size:16px;font-weight:700;color:#1e293b;}
.mval.pos{color:#16a34a;} .mval.neg{color:#dc2626;}
.section-hdr{font-size:16px;font-weight:600;color:#1e293b;margin:28px 0 16px;padding-bottom:10px;border-bottom:1px solid #f1f5f9;}
.dl-card{background:white;border:1px solid #e2e8f0;border-radius:12px;padding:20px;text-align:center;}
.dl-icon{font-size:28px;margin-bottom:8px;}
.dl-title{font-size:13px;font-weight:600;color:#1e293b;margin-bottom:4px;}
.dl-desc{font-size:11px;color:#94a3b8;margin-bottom:14px;line-height:1.4;}
div[data-testid="stButton"] button[kind="primary"]{background:#6366f1;border:none;border-radius:8px;font-weight:600;font-size:14px;}
div[data-testid="stButton"] button[kind="primary"]:hover{background:#4f46e5;}
.stProgress > div > div > div > div{background-color:#6366f1;}
</style>
""", unsafe_allow_html=True)

def fmt(v, c="$"):
    if v is None: return "—"
    if abs(v)>=1e6: return f"{c}{v/1e6:.1f}M"
    if abs(v)>=1e3: return f"{c}{v/1e3:.0f}K"
    return f"{c}{v:,.0f}"

def fpct(v):
    return "—" if v is None else f"{v:.1%}"

def cc(v):
    return "" if v is None else ("pos" if v>=0 else "neg")

def make_zip(pfiles, mbytes):
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as z:
        for n,d in pfiles.items(): z.writestr(f"processed/{n}",d)
        z.writestr("model/forecast_model.xlsx",mbytes)
    return buf.getvalue()

# ── Hero
st.markdown("""
<div class="hero">
  <div class="hero-badge">✦ Powered by Claude AI</div>
  <h1>Financial Forecasting Tool</h1>
  <p>Upload financial statement PDFs for any company. Claude extracts the data, we build a forecast model — ready to download in seconds.</p>
</div>
""", unsafe_allow_html=True)

# ── Steps
st.markdown("""
<div class="steps-row">
  <div class="step-card"><div class="step-num active">1</div><div class="step-title">Upload PDFs</div><div class="step-desc">Annual reports or quarterly statements — any company</div></div>
  <div class="step-card"><div class="step-num">2</div><div class="step-title">AI Extraction</div><div class="step-desc">Claude reads each statement and pulls all financial data</div></div>
  <div class="step-card"><div class="step-num">3</div><div class="step-title">Download Model</div><div class="step-desc">Formula-driven Excel with history, drivers, forecast &amp; scenarios</div></div>
</div>
""", unsafe_allow_html=True)

# ── Inputs
col_l, col_r = st.columns([3,2], gap="large")

with col_l:
    st.markdown('<div class="card"><h3>Financial Statements</h3><p>Upload one PDF per reporting period. Multiple years recommended.</p></div>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader("", type=["pdf"], accept_multiple_files=True, label_visibility="collapsed")

with col_r:
    st.markdown('<div class="card"><h3>Anthropic API Key</h3><p>Get yours at console.anthropic.com</p></div>', unsafe_allow_html=True)
    api_key = st.text_input("", type="password", value=os.environ.get("ANTHROPIC_API_KEY",""), placeholder="sk-ant-...", label_visibility="collapsed")
    model   = st.selectbox("Model", ["claude-sonnet-4-6","claude-opus-4-6","claude-haiku-4-5-20251001"], index=0)
    if uploaded_files:
        items = "".join(f'<div style="font-size:12px;color:#64748b;padding:2px 0">📄 {f.name}</div>' for f in uploaded_files)
        st.markdown(f'<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px;margin-top:8px"><div style="font-size:13px;font-weight:600;color:#1e293b;margin-bottom:6px">{len(uploaded_files)} file(s) ready</div>{items}</div>', unsafe_allow_html=True)

if not uploaded_files:
    st.info("⬆️  Upload at least one PDF to get started.")
    st.stop()
if not api_key:
    st.warning("Enter your Anthropic API key to continue.")
    st.stop()

_, bc, _ = st.columns([1,2,1])
with bc:
    run = st.button("Extract & Build Forecast Model →", type="primary", use_container_width=True)
if not run: st.stop()

# ── Extract
st.markdown('<div class="section-hdr">Extracting Data</div>', unsafe_allow_html=True)
results, period_files, errors = [], {}, []
prog = st.progress(0)
stat = st.empty()

for i, f in enumerate(uploaded_files):
    stat.markdown(f'<div style="font-size:13px;color:#64748b">Processing <b>{f.name}</b>… ({i+1} of {len(uploaded_files)})</div>', unsafe_allow_html=True)
    with st.spinner(""):
        raw = extract_financials(f.read(), api_key, model)
        raw["_source_filename"] = f.name
        data = normalize_extracted_data(raw)
    if "_extraction_error" in data:
        errors.append(f.name)
        st.markdown(f'<div class="err">✕ <b>{f.name}</b> — {data["_extraction_error"]}</div>', unsafe_allow_html=True)
    else:
        results.append(data)
        cur = data.get("currency","$")
        st.markdown(f'<div class="ok">✓ <b>{f.name}</b> · <b>{data.get("company_name","")}</b> · <b>{data.get("period","")}</b> · Revenue {fmt(data.get("revenue"),cur)} · Net Income {fmt(data.get("net_income"),cur)} · GM {fpct(data.get("gross_margin_pct"))}</div>', unsafe_allow_html=True)
        period_files[f'{data["period"]}_financials.xlsx'] = build_period_file(data)
    prog.progress((i+1)/len(uploaded_files))

stat.empty()
if not results:
    st.error("No files processed. Check your API key and try again.")
    st.stop()

# ── Dashboard
results = sorted(results, key=lambda d: d.get("period",""))
st.markdown('<div class="section-hdr">Results</div>', unsafe_allow_html=True)

cols = st.columns(len(results))
for i,d in enumerate(results):
    with cols[i]:
        cur = d.get("currency","$")
        per = d.get("period","")
        rev = d.get("revenue"); ni = d.get("net_income"); gm = d.get("gross_margin_pct")
        ebit = d.get("operating_income"); ocf = d.get("ocf")
        st.metric(f"Revenue — {per}", fmt(rev,cur))
        st.markdown(f"""
<div class="metric-inner">
  <div class="metric-per">{per}</div>
  <div class="mrow"><span class="mlabel">Gross Margin</span><span class="mval">{fpct(gm)}</span></div>
  <div class="mrow"><span class="mlabel">EBIT</span><span class="mval {cc(ebit)}">{fmt(ebit,cur)}</span></div>
  <div class="mrow"><span class="mlabel">Net Income</span><span class="mval {cc(ni)}">{fmt(ni,cur)}</span></div>
  <div class="mrow"><span class="mlabel">Op. Cash Flow</span><span class="mval {cc(ocf)}">{fmt(ocf,cur)}</span></div>
</div>""", unsafe_allow_html=True)

# ── Charts
if len(results) > 1:
    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
    labels = [d.get("period","") for d in results]
    cur = results[0].get("currency","$")
    c1, c2 = st.columns(2, gap="large")

    with c1:
        fig = go.Figure()
        fig.add_bar(name="COGS",         x=labels, y=[d.get("cogs") for d in results],         marker_color="#e2e8f0")
        fig.add_bar(name="Gross Profit", x=labels, y=[d.get("gross_profit") for d in results], marker_color="#6366f1")
        fig.update_layout(title=dict(text="Revenue Breakdown",font=dict(size=14,family="Inter",color="#1e293b"),x=0),
            barmode="stack",paper_bgcolor="white",plot_bgcolor="white",font=dict(family="Inter",size=11,color="#64748b"),
            legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1,font=dict(size=11)),
            margin=dict(l=0,r=0,t=40,b=0),height=280,
            yaxis=dict(tickprefix=cur,gridcolor="#f1f5f9",tickfont=dict(size=10)),
            xaxis=dict(tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        gm_vals   = [(d.get("gross_margin_pct") or 0)*100 for d in results]
        ebit_vals = [((d.get("operating_income") or 0)/max(d.get("revenue") or 1,1))*100 for d in results]
        ni_vals   = [((d.get("net_income") or 0)/max(d.get("revenue") or 1,1))*100 for d in results]
        fig2 = go.Figure()
        fig2.add_scatter(name="Gross Margin",x=labels,y=gm_vals,  mode="lines+markers",line=dict(color="#6366f1",width=2.5),marker=dict(size=7))
        fig2.add_scatter(name="EBIT Margin", x=labels,y=ebit_vals,mode="lines+markers",line=dict(color="#f59e0b",width=2.5,dash="dot"),marker=dict(size=7))
        fig2.add_scatter(name="Net Margin",  x=labels,y=ni_vals,  mode="lines+markers",line=dict(color="#10b981",width=2.5,dash="dash"),marker=dict(size=7))
        fig2.add_hline(y=0,line_color="#e2e8f0",line_width=1)
        fig2.update_layout(title=dict(text="Margin Trends",font=dict(size=14,family="Inter",color="#1e293b"),x=0),
            paper_bgcolor="white",plot_bgcolor="white",font=dict(family="Inter",size=11,color="#64748b"),
            legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1,font=dict(size=11)),
            margin=dict(l=0,r=0,t=40,b=0),height=280,
            yaxis=dict(ticksuffix="%",gridcolor="#f1f5f9",tickfont=dict(size=10)),
            xaxis=dict(tickfont=dict(size=11)))
        st.plotly_chart(fig2, use_container_width=True)

# ── Build & Download
with st.spinner("Building forecast model..."):
    model_bytes = build_forecast_model(results)

st.markdown('<div class="section-hdr">Download</div>', unsafe_allow_html=True)
d1, d2, d3 = st.columns(3, gap="medium")

with d1:
    st.markdown('<div class="dl-card"><div class="dl-icon">📄</div><div class="dl-title">Per-Period Files</div><div class="dl-desc">Raw extract + structured data for each uploaded period</div></div>', unsafe_allow_html=True)
    for fname, fbytes in period_files.items():
        st.download_button(f"↓ {fname}", data=fbytes, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

with d2:
    st.markdown('<div class="dl-card"><div class="dl-icon">📊</div><div class="dl-title">Forecast Model</div><div class="dl-desc">Historical model, drivers, forecast &amp; scenario analysis</div></div>', unsafe_allow_html=True)
    st.download_button("↓ forecast_model.xlsx", data=model_bytes, file_name="forecast_model.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

with d3:
    st.markdown('<div class="dl-card"><div class="dl-icon">🗜️</div><div class="dl-title">Everything</div><div class="dl-desc">All files in a single ZIP, organized by folder</div></div>', unsafe_allow_html=True)
    st.download_button("↓ financial_forecast.zip", data=make_zip(period_files,model_bytes),
        file_name="financial_forecast.zip", mime="application/zip", use_container_width=True)

# ── Raw data
with st.expander("View extracted data"):
    for d in results:
        st.markdown(f"**{d.get('company_name','')} — {d.get('period','')}**")
        c1,c2,c3 = st.columns(3)
        cur = d.get("currency","$")
        with c1:
            st.caption("Income Statement")
            st.dataframe({"Line Item":["Revenue","COGS","Gross Profit","Total OPEX","EBIT","Net Income"],
                "Value":[fmt(d.get(k),cur) for k in ["revenue","cogs","gross_profit","total_opex","operating_income","net_income"]]},
                hide_index=True, use_container_width=True)
        with c2:
            st.caption("Cash Flow")
            st.dataframe({"Line Item":["Op. Cash Flow","Capex","Investing CF","Financing CF"],
                "Value":[fmt(d.get(k),cur) for k in ["ocf","capex","investing_cf","financing_cf"]]},
                hide_index=True, use_container_width=True)
        with c3:
            st.caption("Balance Sheet")
            st.dataframe({"Line Item":["Cash","AR","Inventory","PP&E","Total Assets","Total Equity"],
                "Value":[fmt(d.get(k),cur) for k in ["cash","accounts_receivable","inventories","capital_assets","total_assets","total_equity"]]},
                hide_index=True, use_container_width=True)
        st.divider()
