from pathlib import Path
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"


def init_shared_context() -> None:
    """Initialize values available to both application pages.

    Put cross-page outputs in st.session_state["shared"] so one application can
    pass parsed IDs, selected XML bytes, validity, or other results to the other.
    """
    if "shared" not in st.session_state:
        st.session_state["shared"] = {
            "case_id": "",
            "source_xml_bytes": None,
            "processed_xml_bytes": None,
            "triage_rows": [],
            "quality_review_summary": {},
        }


@st.cache_data(show_spinner=False)
def read_shared_file_bytes(relative_name: str) -> bytes:
    """Read a common repository file once per Streamlit cache lifecycle."""
    path = DATA_DIR / relative_name
    return path.read_bytes()


def set_shared_value(key: str, value) -> None:
    init_shared_context()
    st.session_state["shared"][key] = value


def get_shared_value(key: str, default=None):
    init_shared_context()
    return st.session_state["shared"].get(key, default)
