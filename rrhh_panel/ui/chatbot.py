from __future__ import annotations

import streamlit as st

from rrhh_panel.ai.prompts import SYSTEM_PROMPT
from rrhh_panel.ai.client import chat_stream


def render_chat(context_summary: str = "") -> None:
    if "chat_messages" not in st.session_state:
        st.session_state.chat_messages = [{"role": "system", "content": SYSTEM_PROMPT}]
        if context_summary.strip():
            st.session_state.chat_messages.append(
                {"role": "system", "content": f"Contexto (resumen):\n{context_summary.strip()}"}
            )

    # Render historial (no mostrar system)
    for m in st.session_state.chat_messages:
        if m["role"] in ("user", "assistant"):
            with st.chat_message(m["role"]):
                st.markdown(m["content"])

    user_text = st.chat_input("Pregunta qué significa algo del panel…")
    if not user_text:
        return

    st.session_state.chat_messages.append({"role": "user", "content": user_text})
    with st.chat_message("user"):
        st.markdown(user_text)

    with st.chat_message("assistant"):
        placeholder = st.empty()
        acc = ""

        try:
            resp = chat_stream(
                messages=st.session_state.chat_messages,
                model="glm-4.7-flash",
                stream=True,
                temperature=0.2,
                max_tokens=1200,
            )
            for chunk in resp:
                delta = None
                try:
                    delta = chunk.choices[0].delta.content
                except Exception:
                    delta = None

                if delta:
                    acc += delta
                    placeholder.markdown(acc)

        except Exception as e:
            placeholder.error(f"Error llamando al modelo: {e}")
            return

    st.session_state.chat_messages.append({"role": "assistant", "content": acc})
