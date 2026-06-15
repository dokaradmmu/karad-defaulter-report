import streamlit as st
import requests
import io
from datetime import date, timedelta
from report_builder import build_report

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Karad Division — Defaulter Report",
    page_icon="📮",
    layout="wide",
)

MASTER_RAW_URL = (
    "https://raw.githubusercontent.com/dokaradmmu/karad-defaulter-report/main/Office_Master_File.xlsx"
)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(
    """
    <div style="background:#1F3864;padding:18px 24px;border-radius:8px;margin-bottom:24px">
        <h2 style="color:white;margin:0;font-family:Arial">📮 Karad Division — Daily Defaulter Report</h2>
        <p style="color:#DCE6F1;margin:4px 0 0 0;font-size:14px">
            Office of the Superintendent of Post Offices, Karad Division
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Sidebar — master file management ─────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Master File")
    st.caption("Office_Master_File.xlsx is embedded in the GitHub repo and loaded automatically each run.")

    uploaded_master = st.file_uploader(
        "Upload new master file (replaces repo copy)",
        type=["xlsx"],
        key="master_upload",
        help="Only needed when office structure changes. Push the new file to the repo root."
    )

    if uploaded_master:
        st.success("New master file loaded for this session. Push it to the repo to make it permanent.")

    st.divider()
    st.markdown("**App info**")
    st.caption("Karad Division, Pune Region, Maharashtra Circle")
    st.caption("KPI threshold: 90.00%")
    st.caption("Excluded: Shenawadi B.O, Yeralwadi B.O")

# ── Load master file ──────────────────────────────────────────────────────────
@st.cache_data(ttl=3600, show_spinner=False)
def fetch_master_from_repo():
    resp = requests.get(MASTER_RAW_URL, timeout=15)
    resp.raise_for_status()
    return resp.content

def get_master_bytes():
    if uploaded_master:
        return uploaded_master.read()
    try:
        return fetch_master_from_repo()
    except Exception as e:
        st.error(f"❌ Could not load master file from GitHub: {e}")
        st.stop()

# ── Date inputs ───────────────────────────────────────────────────────────────
today = date.today()
yesterday = today - timedelta(days=1)
first_of_month = today.replace(day=1)

st.markdown("### 📅 Report Dates")
col1, col2, col3 = st.columns(3)
with col1:
    cum_from = st.date_input("Cumulative — From date", value=first_of_month, format="DD/MM/YYYY")
with col2:
    cum_to = st.date_input("Cumulative — To date", value=yesterday, format="DD/MM/YYYY")
with col3:
    daily_date = st.date_input("Daily date", value=yesterday, format="DD/MM/YYYY")

report_date = daily_date + timedelta(days=1)
st.info(
    f"📋 Report date will be: **{report_date.strftime('%d.%m.%Y')}** &nbsp;|&nbsp; "
    f"Filename: **Defaulter_Report_{report_date.strftime('%d_%m_%Y')}.xlsx**"
)

if cum_from > cum_to:
    st.error("❌ Cumulative 'From date' cannot be after 'To date'.")
    st.stop()

# ── File upload slots ─────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("### 📁 Upload CSV Files")
st.caption("File names don't matter — drop each file in its designated slot.")

col_a, col_b = st.columns(2)
col_c, col_d = st.columns(2)

with col_a:
    st.markdown(
        f"**① Cumulative Delivery Productivity**  \n"
        f"<span style='color:#555;font-size:13px'>Period: {cum_from.strftime('%d.%m.%Y')} to {cum_to.strftime('%d.%m.%Y')}</span>",
        unsafe_allow_html=True
    )
    dp_cum_file = st.file_uploader("", type=["csv"], key="dp_cum")

with col_b:
    st.markdown(
        f"**② Daily Delivery Productivity**  \n"
        f"<span style='color:#555;font-size:13px'>Date: {daily_date.strftime('%d.%m.%Y')}</span>",
        unsafe_allow_html=True
    )
    dp_daily_file = st.file_uploader("", type=["csv"], key="dp_daily")

with col_c:
    st.markdown(
        f"**③ Cumulative DSS Usage**  \n"
        f"<span style='color:#555;font-size:13px'>Period: {cum_from.strftime('%d.%m.%Y')} to {cum_to.strftime('%d.%m.%Y')}</span>",
        unsafe_allow_html=True
    )
    dss_cum_file = st.file_uploader("", type=["csv"], key="dss_cum")

with col_d:
    st.markdown(
        f"**④ Daily DSS Usage**  \n"
        f"<span style='color:#555;font-size:13px'>Date: {daily_date.strftime('%d.%m.%Y')}</span>",
        unsafe_allow_html=True
    )
    dss_daily_file = st.file_uploader("", type=["csv"], key="dss_daily")

# ── Generate ──────────────────────────────────────────────────────────────────
st.markdown("---")

all_uploaded = all([dp_cum_file, dp_daily_file, dss_cum_file, dss_daily_file])

if not all_uploaded:
    missing = []
    if not dp_cum_file:   missing.append("① Cumulative DP")
    if not dp_daily_file: missing.append("② Daily DP")
    if not dss_cum_file:  missing.append("③ Cumulative DSS")
    if not dss_daily_file: missing.append("④ Daily DSS")
    st.warning(f"⚠️ Waiting for: {', '.join(missing)}")

generate_btn = st.button(
    "🚀 Generate Defaulter Report",
    disabled=not all_uploaded,
    type="primary",
    use_container_width=True,
)

if generate_btn and all_uploaded:
    with st.spinner("Building report…"):
        try:
            master_bytes = get_master_bytes()

            xlsx_bytes, summary, unmatched, date_tag = build_report(
                master_file=io.BytesIO(master_bytes),
                dp_cum_file=io.BytesIO(dp_cum_file.read()),
                dp_daily_file=io.BytesIO(dp_daily_file.read()),
                dss_cum_file=io.BytesIO(dss_cum_file.read()),
                dss_daily_file=io.BytesIO(dss_daily_file.read()),
                cum_from=cum_from,
                cum_to=cum_to,
                daily_date=daily_date,
            )

            # ── Unmatched warning ──────────────────────────────────────────
            if unmatched:
                st.warning(
                    f"⚠️ **{len(unmatched)} office ID(s) in CSVs not found in master file — excluded from report:**\n\n"
                    + ", ".join(unmatched)
                )

            # ── Summary table ──────────────────────────────────────────────
            st.markdown("### ✅ Report Generated — Summary")
            total_dp = sum(v["dp"] for v in summary.values())
            total_dss = sum(v["dss"] for v in summary.values())

            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("Karad East — DP Defaulters",  summary["Karad East"]["dp"])
            col2.metric("Karad East — DSS Defaulters", summary["Karad East"]["dss"])
            col3.metric("Vaduj — DP Defaulters",        summary["Vaduj"]["dp"])
            col4.metric("Vaduj — DSS Defaulters",       summary["Vaduj"]["dss"])
            col5.metric("Karad West — DP Defaulters",  summary["Karad West"]["dp"])

            col6, col7, col8 = st.columns([1, 1, 3])
            col6.metric("Karad West — DSS Defaulters", summary["Karad West"]["dss"])
            col7.metric("Total DP Defaulters",  total_dp)
            col8.metric("Total DSS Defaulters", total_dss)

            # ── Download ───────────────────────────────────────────────────
            filename = f"Defaulter_Report_{date_tag}.xlsx"
            st.download_button(
                label=f"⬇️ Download {filename}",
                data=xlsx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"❌ Error generating report: {e}")
            st.exception(e)
