import streamlit as st

st.set_page_config(
    page_title="Chapman Tools",
    page_icon="🚗",
    layout="centered"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700&family=Barlow:wght@400;500&display=swap');
    html, body, [class*="css"] { font-family: 'Barlow', sans-serif; }
    h1, h2, h3 { font-family: 'Barlow Condensed', sans-serif !important; letter-spacing: 0.5px; }
    .tool-card {
        background: white;
        border: 1px solid #e0e0e0;
        border-left: 5px solid #003399;
        border-radius: 6px;
        padding: 20px 24px;
        margin: 10px 0;
        cursor: pointer;
        transition: box-shadow 0.2s;
    }
    .tool-card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
    .tool-title { font-family: 'Barlow Condensed', sans-serif; font-size: 22px; font-weight: 700; color: #003399; }
    .tool-desc  { font-size: 14px; color: #555; margin-top: 4px; }
    .hero { text-align: center; padding: 32px 0 24px; }
    .hero h1 { font-family: 'Barlow Condensed', sans-serif; font-size: 48px; font-weight: 700; color: #003399; letter-spacing: 1px; }
    .hero p  { font-size: 16px; color: #666; margin-top: 6px; }
    [data-testid="stPageLink"] a {
        background-color: #003399 !important;
        color: white !important;
        border-radius: 4px !important;
        font-family: 'Barlow Condensed', sans-serif !important;
        font-size: 16px !important;
        font-weight: 600 !important;
        letter-spacing: 0.5px !important;
        padding: 8px 16px !important;
        text-decoration: none !important;
        display: block !important;
        text-align: center !important;
    }
    [data-testid="stPageLink"] a:hover { background-color: #002277 !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <h1>🚗 Chapman Tools</h1>
    <p>Internal tools to make your job easier — select a tool from the sidebar or below.</p>
</div>
""", unsafe_allow_html=True)

st.divider()

st.markdown("### Available Tools")

tools = [
    ("📄", "Invoice Splitter",      "Upload a combined invoice PDF and split it into individual files with an Excel summary.",    "pages/1_Invoice_Splitter.py"),
    ("🏷️", "Loaner Label Generator", "Upload your vehicle list and generate print-ready Avery 8460 loaner labels with QR codes.", "pages/2_Loaner_Labels.py"),
]

for icon, title, desc, page in tools:
    st.markdown(f"""
    <div class="tool-card">
        <div class="tool-title">{icon} {title}</div>
        <div class="tool-desc">{desc}</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button(f"Open {title} →", key=title, use_container_width=True):
        st.switch_page(page)
    st.markdown("<div style='margin-bottom:8px'></div>", unsafe_allow_html=True)

st.divider()
st.caption("Chapman Automotive · Internal Use Only")
