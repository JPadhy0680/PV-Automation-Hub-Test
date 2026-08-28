import streamlit as st
from shared_context import init_shared_context

st.set_page_config(page_title="PV Automation Hub", page_icon="🧩", layout="wide", initial_sidebar_state="expanded")
init_shared_context()

pages = {
    "Applications": [
        st.Page("pages/triage_app.py", title="E2B R3 XML Triage", icon="📊", default=True),
        st.Page("pages/quality_reviewer.py", title="XML Quality Reviewer", icon="🔎"),
        st.Page("pages/line_listing_splitter.py", title="Line Listing Splitter", icon="📑"),
    ]
}

navigation = st.navigation(pages, position="sidebar", expanded=True)
with st.sidebar:
    st.divider()
    st.caption("PV Automation Hub")
navigation.run()
