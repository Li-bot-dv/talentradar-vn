"""
Recruiter AI Agent - Premium Interactive UI (Mode A: Google X-Ray Sourcing)
100% Account-Safe. Vietnam-focused. Strict current-role & current-company verification.
"""

import os
import streamlit as st
import pandas as pd
from dotenv import load_dotenv, set_key
from recruiter_engine import build_xray_queries, execute_xray_sourcing, export_mapping_recap_to_excel

ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
load_dotenv(ENV_PATH)

def _get_secret(key_name: str) -> str:
    """Reads API key from local .env / env vars or Streamlit Cloud Secrets."""
    val = os.getenv(key_name, "")
    if val:
        return val
    try:
        return str(st.secrets.get(key_name, ""))
    except Exception:
        return ""

st.set_page_config(
    page_title="TalentRadar VN · AI Sourcing Agent",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =====================================================
# PREMIUM SAAS THEME & STYLING
# =====================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* Clean light canvas */
    .stApp {
        background-color: #F8FAFC;
    }

    /* Hide default Streamlit chrome */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Sidebar polish */
    section[data-testid="stSidebar"] {
        background-color: #0F172A;
        border-right: 1px solid #1E293B;
    }
    section[data-testid="stSidebar"] * {
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] code {
        background-color: #1E293B !important;
        color: #93C5FD !important;
        border: 1px solid #334155 !important;
        border-radius: 5px !important;
        padding: 0.1rem 0.4rem !important;
        font-size: 0.8rem !important;
    }
    section[data-testid="stSidebar"] svg {
        color: #94A3B8 !important;
        fill: #94A3B8 !important;
    }
    section[data-testid="stSidebar"] .stSlider label,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        color: #F8FAFC !important;
        font-weight: 600 !important;
    }
    section[data-testid="stSidebar"] hr {
        border-color: #1E293B !important;
    }
    section[data-testid="stSidebar"] input,
    section[data-testid="stSidebar"] div[data-baseweb="input"],
    section[data-testid="stSidebar"] div[data-baseweb="base-input"],
    section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
        background-color: #1E293B !important;
        color: #F8FAFC !important;
        border-color: #334155 !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stAlert"] {
        background-color: rgba(16, 185, 129, 0.16) !important;
        border: 1px solid rgba(52, 211, 153, 0.35) !important;
        border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stAlert"] * {
        color: #6EE7B7 !important;
    }
    section[data-testid="stSidebar"] button {
        background-color: #1E293B !important;
        color: #F8FAFC !important;
        border: 1px solid #334155 !important;
        border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] button:hover {
        background-color: #334155 !important;
        border-color: #475569 !important;
    }

    /* Hero header card */
    .hero-banner {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 55%, #1E3A8A 100%);
        border-radius: 16px;
        padding: 1.75rem 2rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.18);
        border: 1px solid rgba(255, 255, 255, 0.08);
    }
    .hero-top {
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 0.75rem;
        margin-bottom: 0.45rem;
    }
    .hero-title {
        font-size: 1.85rem;
        font-weight: 700;
        color: #FFFFFF;
        letter-spacing: -0.02em;
        margin: 0;
    }
    .hero-subtitle {
        color: #CBD5E1;
        font-size: 0.96rem;
        margin: 0;
        line-height: 1.5;
    }
    .safe-pill {
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
        padding: 0.32rem 0.85rem;
        border-radius: 9999px;
        background: rgba(16, 185, 129, 0.16);
        border: 1px solid rgba(52, 211, 153, 0.35);
        color: #6EE7B7;
        font-weight: 600;
        font-size: 0.8rem;
        letter-spacing: 0.01em;
    }

    /* Input section cards */
    .input-card-header {
        font-size: 0.95rem;
        font-weight: 600;
        color: #0F172A;
        margin-bottom: 0.15rem;
        display: flex;
        align-items: center;
        gap: 0.4rem;
    }
    .input-card-hint {
        font-size: 0.8rem;
        color: #64748B;
        margin-bottom: 0.5rem;
    }

    /* Styled TextAreas */
    .stTextArea textarea {
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 10px !important;
        font-size: 0.9rem !important;
        color: #0F172A !important;
        padding: 0.75rem !important;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04) !important;
        transition: border-color 0.15s ease, box-shadow 0.15s ease !important;
    }
    .stTextArea textarea:focus {
        border-color: #2563EB !important;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12) !important;
    }

    /* Scope summary bar */
    .scope-bar {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-left: 4px solid #2563EB;
        border-radius: 10px;
        padding: 0.85rem 1.15rem;
        font-size: 0.88rem;
        color: #334155;
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 0.5rem;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
    }
    .scope-chip {
        background: #EFF6FF;
        color: #1D4ED8;
        font-weight: 600;
        padding: 0.15rem 0.55rem;
        border-radius: 6px;
        font-size: 0.82rem;
    }

    /* Primary CTA button */
    div.stButton > button[kind="primary"],
    div.stDownloadButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.65rem 1.25rem !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25) !important;
        transition: transform 0.1s ease, box-shadow 0.15s ease !important;
    }
    div.stButton > button[kind="primary"]:hover,
    div.stDownloadButton > button[kind="primary"]:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.35) !important;
    }

    /* KPI metric cards */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 1rem;
        margin: 0.75rem 0 1.25rem 0;
    }
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.1rem 1.25rem;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
    }
    .kpi-label {
        font-size: 0.78rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #64748B;
        margin-bottom: 0.35rem;
    }
    .kpi-value {
        font-size: 1.75rem;
        font-weight: 700;
        color: #0F172A;
        line-height: 1.1;
    }
    .kpi-accent-green { color: #059669; }
    .kpi-accent-red { color: #DC2626; }
    .kpi-accent-blue { color: #2563EB; }

    /* Tabs polish */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem;
        border-bottom: 1px solid #E2E8F0;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 0.55rem 1.1rem;
        font-weight: 600;
        font-size: 0.9rem;
        color: #475569;
    }
    .stTabs [aria-selected="true"] {
        color: #1D4ED8 !important;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================
# HERO HEADER
# =====================================================
st.markdown("""
<div class="hero-banner">
    <div class="hero-top">
        <div class="hero-title">🎯 TalentRadar VN · AI Sourcing Agent</div>
        <div class="safe-pill">● Mode A · 100% Account-Safe X-Ray · Vietnam</div>
    </div>
    <p class="hero-subtitle">
        Precision company-by-company talent mapping on LinkedIn Vietnam. Verifies current company &amp; current role automatically — zero LinkedIn login required.
    </p>
</div>
""", unsafe_allow_html=True)

# State
if "candidates" not in st.session_state:
    st.session_state.candidates = []
if "company_status" not in st.session_state:
    st.session_state.company_status = []
if "excel_file" not in st.session_state:
    st.session_state.excel_file = None

# =====================================================
# SIDEBAR CONTROLS
# =====================================================
with st.sidebar:
    st.markdown("### ⚡ Search Engine")
    saved_serper = _get_secret("SERPER_API_KEY")
    saved_google_key = _get_secret("GOOGLE_CSE_API_KEY")
    saved_google_cx = _get_secret("GOOGLE_CSE_CX")
    default_engine_idx = 1 if saved_serper else (2 if (saved_google_key and saved_google_cx) else 0)

    engine_choice = st.selectbox(
        "Engine Mode",
        [
            "Browser Mode (Free · Playwright)",
            "Serper.dev API (Fast · 2,500 Free)",
            "Google Custom Search API (100/Day)",
        ],
        index=default_engine_idx,
        label_visibility="collapsed",
        help="Browser Mode is free & local only. Serper.dev and Google CSE work on Streamlit Cloud with zero rate-limits."
    )

    search_provider = "browser"
    api_key = ""
    cse_cx = ""

    if "Serper" in engine_choice:
        search_provider = "serper"
        api_key = st.text_input(
            "Serper.dev API Key",
            value=saved_serper,
            type="password",
            placeholder="Paste key from serper.dev",
        )
        if st.button("💾 Save Serper Key", use_container_width=True):
            if api_key.strip():
                if not os.path.exists(ENV_PATH):
                    open(ENV_PATH, "a").close()
                set_key(ENV_PATH, "SERPER_API_KEY", api_key.strip())
                st.success("Saved to .env!")
            else:
                st.warning("Enter a key first.")
        st.caption("Get 2,500 free searches at [serper.dev](https://serper.dev) (no credit card).")

    elif "Google Custom" in engine_choice:
        search_provider = "google_cse"
        api_key = st.text_input(
            "Google API Key",
            value=saved_google_key,
            type="password",
            placeholder="AIza...",
        )
        cse_cx = st.text_input(
            "Search Engine ID (CX)",
            value=saved_google_cx,
            placeholder="e.g. a1b2c3d4e5f6g7h8i",
        )
        if st.button("💾 Save Google CSE Keys", use_container_width=True):
            if api_key.strip() and cse_cx.strip():
                if not os.path.exists(ENV_PATH):
                    open(ENV_PATH, "a").close()
                set_key(ENV_PATH, "GOOGLE_CSE_API_KEY", api_key.strip())
                set_key(ENV_PATH, "GOOGLE_CSE_CX", cse_cx.strip())
                st.success("Saved to .env!")
            else:
                st.warning("Enter both API Key and CX ID.")
        st.caption("100 free searches/day, refreshes daily forever.")

    st.divider()
    st.markdown("### ⚙️ Search Configuration")
    max_per_comp = st.slider(
        "Max Candidates per Company",
        min_value=2,
        max_value=50,
        value=10,
        help="Maximum verified profiles to collect per target company."
    )
    max_pages = st.slider(
        "Google Pages to Scrape per Query",
        min_value=1,
        max_value=10,
        value=3,
        help="1 = fastest, 10 = deepest coverage (Pages 1–10). Stops early automatically when a page yields 0 matches or Google runs out of results."
    )
    strict_headline_bg = st.toggle(
        "Strict Role Match (Title × Background)",
        value=True,
        help="Require BOTH Title and Background in the candidate's current role (e.g., Talent Acquisition + Manager). Turn OFF for technical stack keywords that live in the profile body."
    )

    st.divider()
    st.markdown("### 📍 Target Market")
    st.markdown("**Vietnam** (`vn.linkedin.com/in`)")
    st.caption("Automatically filters out ex-employees, job postings, and unrelated roles.")

    st.divider()
    st.markdown("### 💡 Quick Guide")
    st.markdown("""
    - **Companies**: 1 per line or comma-separated
    - **Titles**: Treated as `OR` (e.g. `Manager`, `Head`)
    - **Background**: Treated as `OR` (e.g. `Treasury`, `ALM`, `ALCM`)
    """)

# =====================================================
# INTAKE CARDS
# =====================================================
col1, col2, col3 = st.columns(3, gap="medium")

with col1:
    st.markdown('<div class="input-card-header">🏢 1. Target Companies</div>', unsafe_allow_html=True)
    st.markdown('<div class="input-card-hint">One per line or comma-separated</div>', unsafe_allow_html=True)
    companies_raw = st.text_area(
        "Target Companies",
        value="MoMo\nVNPay\nZaloPay\nShopee\nGrab\nHome Credit",
        height=165,
        label_visibility="collapsed"
    )

with col2:
    st.markdown('<div class="input-card-header">💼 2. Relevant Titles</div>', unsafe_allow_html=True)
    st.markdown('<div class="input-card-hint">Seniority / level words (e.g. Manager, Head, Lead, Specialist)</div>', unsafe_allow_html=True)
    titles_raw = st.text_area(
        "Relevant Titles",
        value="Manager\nHead",
        height=165,
        label_visibility="collapsed"
    )

with col3:
    st.markdown('<div class="input-card-header">🛠️ 3. Background & Skills</div>', unsafe_allow_html=True)
    st.markdown('<div class="input-card-hint">Function or domain keywords (e.g. Talent Acquisition, Legal)</div>', unsafe_allow_html=True)
    bg_raw = st.text_area(
        "Background & Skills",
        value="Talent Acquisition",
        height=165,
        label_visibility="collapsed"
    )

def parse_list(raw: str):
    items = []
    for line in raw.split("\n"):
        for part in line.split(","):
            val = part.strip()
            if val and val not in items:
                items.append(val)
    return items

companies = parse_list(companies_raw)
titles = parse_list(titles_raw)
backgrounds = parse_list(bg_raw)

# Query Preview
queries = build_xray_queries(companies, titles, backgrounds, strict_headline_bg=strict_headline_bg)
with st.expander(f"🔍 Preview X-Ray Queries ({len(companies)} companies · up to {max_pages} pages/query)"):
    for comp, q in queries.items():
        st.markdown(f"**{comp}**")
        st.code(q, language="text")

st.markdown("<div style='height: 0.4rem'></div>", unsafe_allow_html=True)

# =====================================================
# ACTION BAR & SCOPE SUMMARY
# =====================================================
start_col, info_col = st.columns([1, 2.4], gap="medium")
with start_col:
    start_btn = st.button("🚀 Start Automated Sourcing", type="primary", use_container_width=True)

with info_col:
    title_txt = " OR ".join([f'"{t}"' for t in titles]) if titles else "Any Title"
    bg_txt = " OR ".join([f'"{b}"' for b in backgrounds]) if backgrounds else "Any Background"
    engine_badge = {"browser": "Browser", "serper": "Serper.dev API", "google_cse": "Google CSE API"}[search_provider]
    st.markdown(
        f"""
        <div class="scope-bar">
            <span class="scope-chip">{len(companies)} Companies</span>
            <span class="scope-chip">{engine_badge}</span>
            <span>Role: <b>({title_txt})</b> AND <b>({bg_txt})</b></span>
            <span style="color:#94A3B8;">|</span>
            <span>Max <b>{max_per_comp}</b>/company · Up to <b>{max_pages}</b> pages</span>
        </div>
        """,
        unsafe_allow_html=True
    )

if start_btn:
    if not companies or (not titles and not backgrounds):
        st.error("Please provide at least one target company and one title or background keyword.")
    elif search_provider == "serper" and not api_key.strip():
        st.error("Please enter your Serper.dev API Key in the sidebar (or switch back to Browser Mode).")
    elif search_provider == "google_cse" and (not api_key.strip() or not cse_cx.strip()):
        st.error("Please enter both your Google API Key and Search Engine ID (CX) in the sidebar.")
    else:
        status_box = st.empty()
        progress_bar = st.progress(0)

        def update_status(msg: str):
            status_box.info(f"⏳ {msg}")

        try:
            with st.spinner("Mapping Vietnam talent company-by-company..."):
                candidates, comp_status = execute_xray_sourcing(
                    companies=companies,
                    titles=titles,
                    backgrounds=backgrounds,
                    max_per_company=max_per_comp,
                    max_pages=max_pages,
                    strict_headline_bg=strict_headline_bg,
                    search_provider=search_provider,
                    api_key=api_key,
                    cse_cx=cse_cx,
                    status_callback=update_status
                )
                st.session_state.candidates = candidates
                st.session_state.company_status = comp_status

                out_file = "Recruiter_Vietnam_Sourcing_Recap.xlsx"
                export_mapping_recap_to_excel(candidates, comp_status, out_file)
                st.session_state.excel_file = out_file

                progress_bar.progress(100)
                status_box.success(f"✅ Done! Sourced {len(candidates)} verified profiles across {len(companies)} companies.")
                st.rerun()
        except Exception as e:
            status_box.error(f"❌ {str(e)}")

# =====================================================
# RESULTS DASHBOARD
# =====================================================
if st.session_state.candidates or st.session_state.company_status:
    st.markdown("<div style='height: 0.75rem'></div>", unsafe_allow_html=True)
    st.markdown("### 📊 Sourcing Results")

    total = len(st.session_state.candidates)
    total_comps = len(st.session_state.company_status)
    found = sum(1 for c in st.session_state.company_status if c.get("status") == "Found")
    not_found = sum(1 for c in st.session_state.company_status if c.get("status") == "Not Found")
    coverage_pct = int((found / total_comps) * 100) if total_comps > 0 else 0

    st.markdown(f"""
    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-label">Verified Candidates</div>
            <div class="kpi-value kpi-accent-blue">{total}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Companies Found</div>
            <div class="kpi-value kpi-accent-green">{found}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">No Results</div>
            <div class="kpi-value kpi-accent-red">{not_found}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Market Coverage</div>
            <div class="kpi-value">{coverage_pct}%</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    tab_cand, tab_cov = st.tabs([f"📋 Sourced Candidates ({total})", f"📈 Company Coverage ({found}/{total_comps})"])

    with tab_cand:
        if st.session_state.candidates:
            df = pd.DataFrame(st.session_state.candidates)
            df.columns = ["Name", "Title", "Industry / Background", "Company", "LinkedIn Link"]
            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "LinkedIn Link": st.column_config.LinkColumn(
                        "LinkedIn Profile",
                        display_text="Open Profile ↗"
                    )
                }
            )
        else:
            st.warning("No candidates matched all strict criteria. Try broadening titles or turning off Strict Role Match.")

    with tab_cov:
        if st.session_state.company_status:
            df_s = pd.DataFrame(st.session_state.company_status)[["company", "status", "count"]]
            df_s.columns = ["Company", "Status", "Profiles Found"]

            def color_status(val):
                return "background-color: #D1FAE5; color: #065F46; font-weight: 600;" if val == "Found" \
                    else "background-color: #FEE2E2; color: #991B1B; font-weight: 600;"

            st.dataframe(
                df_s.style.map(color_status, subset=["Status"]),
                use_container_width=True,
                hide_index=True
            )

    st.markdown("<div style='height: 0.5rem'></div>", unsafe_allow_html=True)
    if st.session_state.excel_file and os.path.exists(st.session_state.excel_file):
        with open(st.session_state.excel_file, "rb") as f:
            st.download_button(
                label="📥 Download Formatted Excel Recap (.xlsx) — 2 Sheets",
                data=f.read(),
                file_name="Recruiter_Vietnam_Sourcing_Recap.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
