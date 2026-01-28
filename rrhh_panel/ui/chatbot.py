from __future__ import annotations
import streamlit as st
from rrhh_panel.ai.prompts import SYSTEM_PROMPT
from rrhh_panel.ai.client import chat_stream

def render_chat(context_summary: str = "") -> None:
    if "chat_messages" not in st.session_state:
        st.session_state.chat_messages = [{"role": "system", "content": SYSTEM_PROMPT.strip()}]
        if context_summary.strip():
            st.session_state.chat_messages.append(
                {"role": "system", "content": f"Contexto (resumen):\n{context_summary.strip()}"}
            )

    # historial
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
            resp = chat_stream(st.session_state.chat_messages, model="glm-4.7-flash", stream=True)
            for chunk in resp:
                delta = chunk.choices[0].delta.content if chunk.choices and chunk.choices[0].delta else None
                if delta:
                    acc += delta
                    placeholder.markdown(acc)
        except Exception as e:
            placeholder.error(f"Error llamando al modelo: {e}")
            return

    st.session_state.chat_messages.append({"role": "assistant", "content": acc})
