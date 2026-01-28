# -*- coding: utf-8 -*-
from __future__ import annotations

import streamlit as st

from rrhh_panel.config.texts import (
    APP_TITLE,
    APP_SUBTITLE,
    VIEW_1,
    VIEW_2,
    MSG_NEED_DATA,
    MSG_LOAD_FILE_TO_START,
)
from rrhh_panel.ui.theming import load_css
from rrhh_panel.ui.layout import topbar, columns_layout
from rrhh_panel.ui.sidebar import render_sidebar
from rrhh_panel.ui.dashboard import render_dashboard
from rrhh_panel.ui.descriptives import render_descriptives
from rrhh_panel.ui.floating_chat import render_floating_chat


def _build_context_summary(view: str) -> str:
    """
    Resumen corto para el chatbot.
    NO debe romper si aún no hay datos cargados.
    """
    g = st.session_state.get("__globals__")
    if not g:
        return f"Vista={view} | (sin datos cargados aún)"

    try:
        start_dt = g.get("start_dt")
        end_dt = g.get("end_dt")
        period_label = g.get("period_label", "-")

        if start_dt is not None and end_dt is not None:
            return f"Vista={view} | Periodo={period_label} | Rango={start_dt.date()} a {end_dt.date()}"
        return f"Vista={view} | Periodo={period_label} | (rango no definido)"
    except Exception:
        return f"Vista={view} | (contexto no disponible)"


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    load_css()

    st.title(APP_TITLE)
    st.caption(APP_SUBTITLE)

    # Top bar
    show_filters, view = topbar()

    # Layout principal
    col_filters, col_main = columns_layout(show_filters)

    # Sidebar (carga datos / filtros / opciones)
    if show_filters:
        with col_filters:
            render_sidebar()

    # Chat flotante SIEMPRE visible (no depende de datos)
    context_summary = _build_context_summary(view=view)
    render_floating_chat(context_summary=context_summary)

    # Main content
    with col_main:
        g = st.session_state.get("__globals__")
        fs = st.session_state.get("__fs__")
        opts = st.session_state.get("__opts__")

        if not g or not fs or not opts:
            st.warning(MSG_NEED_DATA if not show_filters else MSG_LOAD_FILE_TO_START)
            return

        if view == VIEW_1:
            render_dashboard(g=g, fs=fs, opts=opts)
        else:
            render_descriptives(g=g, fs=fs, opts=opts)


if __name__ == "__main__":
    main()
