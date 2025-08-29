import requests
import streamlit as st

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_html(url: str) -> tuple[str, bytes]:
    resp = requests.get(
        url,
        timeout=30,
        headers={"User-Agent": "Mozilla/5.0 (compatible; ContentRecTool/1.0)"},
    )
    resp.raise_for_status()
    return resp.url, resp.content
