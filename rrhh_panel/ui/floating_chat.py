from __future__ import annotations

import streamlit as st
from streamlit_float import float_init

from rrhh_panel.ui.chatbot import render_chat


def render_floating_chat(context_summary: str = "") -> None:
    float_init()

    if "chat_open" not in st.session_state:
        st.session_state.chat_open = False

    # Botón flotante (launcher)
    launcher = st.container()
    with launcher:
        if st.button("💬", key="chat_launcher", help="Abrir / cerrar asistente"):
            st.session_state.chat_open = not st.session_state.chat_open
            st.rerun()

    # ABAJO-IZQUIERDA
    launcher.float("bottom: 1.25rem; left: 1.25rem; width: 3.25rem; z-index: 9999;")

    # Si cerrado, salir (evita panel no definido)
    if not st.session_state.chat_open:
        return

    # Panel flotante
    panel = st.container()
    with panel:
        h1, h2, h3 = st.columns([1, 6, 1])
        with h1:
            st.write("💬")
        with h2:
            st.markdown("**Asistente**")
        with h3:
            if st.button("✕", key="chat_close", help="Cerrar"):
                st.session_state.chat_open = False
                st.rerun()

        st.divider()
        render_chat(context_summary=context_summary)

    panel.float(
        "bottom: 5rem; left: 1.25rem; width: 380px; "
        "background-color: rgba(255,255,255,0.98); padding: 0.75rem; "
        "border: 1px solid rgba(0,0,0,0.08); border-radius: 14px; "
        "box-shadow: 0 10px 30px rgba(0,0,0,0.15); z-index: 9999;"
    )
