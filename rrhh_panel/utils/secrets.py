from __future__ import annotations
import os

def get_secret(name: str, default: str = "") -> str:
    try:
        import streamlit as st  # solo si está corriendo en Streamlit
        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:
        pass
    return os.getenv(name, default)
