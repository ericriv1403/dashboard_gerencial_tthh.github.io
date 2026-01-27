# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import os
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# =============================================================================
# 0) TEXTOS / TÍTULOS (EDITA AQUÍ PARA CAMBIAR NOMBRES RÁPIDO)
# =============================================================================
APP_TITLE = "Panel RRHH"
APP_SUBTITLE = "Rotación & Supervivencia | Existencias Actuales | PDE + KPI_COST"

LBL_SHOW_FILTERS = "Mostrar filtros"
LBL_VIEW_PICK = "Vista"
VIEW_1 = "Rotación y Supervivencia"
VIEW_2 = "Existencias Actuales"

# Panel de control
PANEL_TITLE = "Panel de control"
TAB_DATA = "Datos & Periodo"
TAB_FILTERS = "Filtros"
TAB_OPTIONS = "Opciones"
TAB_COST = "Costo KPI"

# Datos & Periodo
LBL_UPLOAD_MAIN = "Sube Excel/CSV (Historia Personal)"
LBL_PATH_MAIN = "O ruta local (opcional)"
LBL_SHEET_MAIN = "Hoja (Historia Personal)"
LBL_COST_AUTO = "Costo Nominal detectado (opcional)"
LBL_USE_COST_SAME = "Usar Costo Nominal de este mismo Excel"
LBL_UPLOAD_COST = "Sube Excel (Costo Nominal)"
LBL_SHEET_COST = "Hoja (Costo Nominal)"
LBL_PATH_COST = "Ruta local Costo Nominal (opcional)"
LBL_RANGE_PRESET = "Atajo de rango"
LBL_RANGE_SLIDER = "Inicio / Fin"
LBL_GROUP_BY = "Agrupar por"
LBL_SNAPSHOT_DATE = "Snapshot existencias (día)"
LBL_TODAY_CUT = "Hoy (corte)"

# Filtros
LBL_FILTERS_HINT = "Deja vacío = no filtra (equivale a TODOS)."
BTN_CLEAR_FILTERS = "Limpiar filtros"
LBL_SEXO = "Sexo"
LBL_AREA_GEN = "Área General"
LBL_AREA = "Área (nombre)"
LBL_CARGO = "Cargo"
LBL_CLAS = "Clasificación"
LBL_TS = "Trabajadora Social"
LBL_EMP = "Empresa"
LBL_NAC = "Nacionalidad"
LBL_LUG = "Lugar Registro"
LBL_REG = "Región Registro"
LBL_TENURE_BUCKET = "Antigüedad (bucket)"
LBL_AGE_BUCKET = "Edad (bucket)"

# Opciones
LBL_OPT_UNIQUE_DAY = "Salidas: contar personas únicas por día"
LBL_OPT_SHOW_LABELS = "Mostrar etiquetas numéricas"
LBL_OPT_HORIZON = "Horizonte H (días) para Supervivencia"
LBL_OPT_H_CHOICE = "Horizonte"
LBL_OPT_H_CUSTOM = "H (días)"
LBL_OPT_SEMAFORO = "Semáforo KPI (ratio vs meta)"
LBL_OPT_GREEN = "Verde si ≤"
LBL_OPT_YELLOW = "Amarillo si ≤"
LBL_OPT_CONT = "Contingencia"
LBL_OPT_TOP_ROWS = "Máx. filas"
LBL_OPT_TOP_COLS = "Máx. columnas"

# Costo KPI
LBL_COST_HINT = "Si no cargaste 'Costo Nominal', esta parte se desactiva sola."
LBL_COST_YELLOW_MIN = "AMARILLO si ratio vs meta ≥"

# Vista 1 layout
V1_ROW1_KPI_TITLE = "KPIs Ejecutivos"
KPI1_TITLE = "Tasa Rotación Ponderada (RateW_1000 TOTAL) — último periodo"
KPI2_TITLE = "Supervivencia (H días) — TOTAL (Kaplan–Meier simple)"
KPI3_TITLE = "Pérdida de Rendimiento — Σ r_pct salidas (último periodo)"

V1_ROW2_LEFT_TITLE = "Tendencia RateW_1000 TOTAL por periodo"
V1_ROW2_RIGHT_TITLE = "Curva de Supervivencia (KM simple)"
V1_ROW3_TITLE = "Desglose por Área y Cargo (último periodo) — Exposure, Exits_w, RateW_1000"

V1_LEGACY_EXPANDER = "PDE + KPI_COST (legacy, no romper)"
V1_LEGACY_PDE_TITLE = "KPI_PDE (legacy)"
V1_LEGACY_COST_TITLE = "KPI_COST (legacy)"

# Vista 2 layout
V2_ROW1_KPI1 = "Total empleados activos (snapshot)"
V2_ROW1_KPI2 = "Rendimiento promedio (avg r_pct)"
V2_ROW1_KPI3 = "Antigüedad promedio (días desde ÚLTIMO ingreso)"

V2_ROW2_LEFT = "Distribución por Clasificación (snapshot)"
V2_ROW2_RIGHT = "Distribución por Edad (bucket) (snapshot)"
V2_ROW3_TITLE = "Tabla de activos (snapshot) — Área/Departamento, Cargo, Antigüedad exacta, r_pct"

# Descargas
DL_HISTORY_XLSX = "Descargar Excel (Historia)"
DL_CURRENT_XLSX = "Descargar Excel (Actual)"
FILE_HISTORY_XLSX = "rrhh_historia.xlsx"
FILE_CURRENT_XLSX = "rrhh_actual.xlsx"

# Mensajes
MSG_NEED_DATA = "Para cargar datos, enciende 'Mostrar filtros' y usa el panel."
MSG_LOAD_FILE_TO_START = "Carga un archivo para iniciar."
MSG_PATH_NOT_FOUND = "La ruta no existe."
MSG_READ_FAIL = "No se pudo leer el archivo:"
MSG_NO_DATA_FOR_VIEW = "No hay datos suficientes con los filtros actuales."


# =============================================================================
# Config Streamlit
# =============================================================================
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)
st.caption(APP_SUBTITLE)


# =============================================================================
# Columnas esperadas (Historia Personal)
# - %R es opcional pero recomendado (si no está, se asume 1.0)
# =============================================================================
REQUIRED_COLS = [
    "Código Personal",
    "Fecha Inicio Evento",
    "Fecha Fin Evento",
    "Fecha Nacimiento",
    "Clasificación",
    "Sexo",
    "TS_Responsable",
    "Empresa",
    "Área Original",
    "Cargo Actual",
    "Nacionalidad",
    "Lugar Registro",
    "Región Registro",
]

R_COL_CANDIDATES = ["%R", "% R", "R", "R%", "Porcentaje R", "PorcentajeR", "Factor R", "FactorR"]

COL_MAP = {
    "Código Personal": "cod",
    "Fecha Inicio Evento": "ini",
    "Fecha Fin Evento": "fin",
    "Fecha Nacimiento": "fnac",
    "Clasificación": "clas_raw",
    "Sexo": "sexo",
    "TS_Responsable": "ts",
    "Empresa": "emp",
    "Área Original": "area_raw",
    "Cargo Actual": "cargo",
    "Nacionalidad": "nac",
    "Lugar Registro": "lug",
    "Región Registro": "reg",
}

MISSING_LABEL = "SIN DATO"


# =============================================================================
# Columnas esperadas (Costo Nominal)
# =============================================================================
COST_REQUIRED = ["Codigo Personal", "Fecha Inicio Corte", "Fecha Fin Corte", "Costo Nominal Diario"]
COST_COL_MAP = {"Codigo Personal": "cod", "Fecha Inicio Corte": "c_ini", "Fecha Fin Corte": "c_fin", "Costo Nominal Diario": "costo"}


# =============================================================================
# Tablas de referencia
# =============================================================================
AREA_REF: Dict[str, Tuple[str, str]] = {
    "ADM": ("ADMINISTRACIÓN", "ADMINISTRACIÓN"),
    "COMPRAS": ("COMPRAS", "ADMINISTRACIÓN"),
    "CONTA": ("CONTABILIDAD", "ADMINISTRACIÓN"),
    "FIN": ("FINANZAS", "ADMINISTRACIÓN"),
    "ING": ("INGENIERÍA", "ADMINISTRACIÓN"),
    "DISTRIBUCION": ("DISTRIBUCIÓN Y TRÁFICO", "ADMINISTRACIÓN"),
    "PROD": ("PRODUCCIÓN", "ADMINISTRACIÓN"),
    "SSO": ("SEGURIDAD Y SALUD OCUPACIONAL", "ADMINISTRACIÓN"),
    "TTHH": ("TALENTO HUMANO", "ADMINISTRACIÓN"),
    "SISTE": ("SISTEMAS", "ADMINISTRACIÓN"),
    "VENT": ("VENTAS", "ADMINISTRACIÓN"),

    "LAB": ("LABORATORIO", "PRODUCCIÓN – PROPAGACIÓN"),
    "A-4": ("SAN JUAN", "PRODUCCIÓN – CAMPO"),
    "CULTIVOS VARIOS": ("CULTIVOS VARIOS", "PRODUCCIÓN – CAMPO"),
    "MH1": ("MONJASHUAICO 1", "PRODUCCIÓN – CAMPO"),
    "MH2": ("MONJASHUAICO 2", "PRODUCCIÓN – CAMPO"),
    "RIEGO": ("RIEGO", "PRODUCCIÓN – CAMPO"),
    "ORN": ("ORNAMENTALES", "PRODUCCIÓN – ORNAMENTALES"),

    "CLS": ("CLASIFICACIÓN", "PRODUCCIÓN – POSCOSECHA"),
    "EMP": ("EMPAQUE", "PRODUCCIÓN – POSCOSECHA"),
    "SB": ("SALA DE BROTE", "PRODUCCIÓN – POSCOSECHA"),

    "PROP": ("PROPAGACIÓN", "PRODUCCIÓN – PROPAGACIÓN"),

    "MANT": ("MANTENIMIENTO", "PRODUCCIÓN – TRANSVERSAL"),
    "BOD": ("BODEGA", "PRODUCCIÓN – TRANSVERSAL"),
    "DRONES": ("OPERACIÓN DE DRONES", "PRODUCCIÓN – TRANSVERSAL"),
    "MONITOREO": ("MONITOREO", "PRODUCCIÓN – TRANSVERSAL"),

    "CHOFER": ("TRANSPORTE INTERNO", "SERVICIOS GENERALES"),
    "SP": ("SERVICIOS PRESTADOS", "SERVICIOS GENERALES"),
    "SRG": ("SRG (SERVICIOS GENERALES)", "SERVICIOS GENERALES"),
    "џPAS": ("PASANTÍA", "SERVICIOS GENERALES"),
    "PAS": ("PASANTÍA", "SERVICIOS GENERALES"),
    "PRACT": ("PRACTICANTES", "SERVICIOS GENERALES"),
}

CLAS_REF: Dict[str, str] = {
    "ADM": "ADMINISTRATIVO",
    "AGR": "TRABAJADOR AGRÍCOLA",
    "CHOFER": "CHOFER",
    "OCAS": "TRABAJADOR AGRÍCOLA OCASIONAL",
    "PAS": "PASANTÍA",
    "PRACT": "PRACTICANTES",
    "SP": "SERVICIOS PRESTADOS",
}


# =============================================================================
# Helpers base (no cambiar nombres internos)
# =============================================================================
def _to_datetime(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce").dt.normalize()

def today_dt() -> pd.Timestamp:
    return pd.Timestamp(date.today())

def excel_weeknum_return_type_1(d: pd.Series) -> pd.Series:
    # Week number like Excel WEEKNUM(date,1) -> weeks start Sunday
    return d.dt.strftime("%U").astype(int) + 1

def week_end_sun_to_sat(d: pd.Series) -> pd.Series:
    # Week start Sunday, end Saturday
    wd = d.dt.weekday  # Mon=0..Sun=6
    days_since_sun = (wd + 1) % 7
    wstart = d - pd.to_timedelta(days_since_sun, unit="D")
    return (wstart + pd.to_timedelta(6, unit="D")).dt.normalize()

def month_end(d: pd.Series) -> pd.Series:
    return (d + pd.offsets.MonthEnd(0)).dt.normalize()

def add_calendar_fields(df: pd.DataFrame, date_col: str) -> pd.DataFrame:
    out = df.copy()
    d = pd.to_datetime(out[date_col], errors="coerce").dt.normalize()

    out["Día"] = d
    out["Año"] = d.dt.year.astype("int64")
    out["Mes"] = d.dt.month.astype("int64")
    out["Semana"] = excel_weeknum_return_type_1(d).astype("int64")

    yy = (out["Año"] % 100).astype(int).astype(str).str.zfill(2)
    ww = out["Semana"].astype(int).astype(str).str.zfill(2)
    mm = out["Mes"].astype(int).astype(str).str.zfill(2)

    out["CodSem"] = (yy + ww).astype(str)  # YYWW
    out["CodMes"] = (yy + mm).astype(str)  # YYMM

    out["FinSemana"] = week_end_sun_to_sat(d)
    out["FinMes"] = month_end(d)
    return out

def _years_offset_date(d: date, years: int) -> date:
    try:
        return d.replace(year=d.year + years)
    except ValueError:
        return d.replace(month=2, day=28, year=d.year + years)

def _norm_cat(s: pd.Series, missing_label: str = MISSING_LABEL) -> pd.Series:
    x = s.copy()
    x = x.replace(["", " ", "None", "nan", "NaN", "NAN", "NaT"], np.nan)
    x = x.fillna(missing_label)
    x = x.astype(str).str.strip()
    x = x.replace({"": missing_label})
    return x

def _safe_table_for_streamlit(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c) for c in out.columns]
    return out

def _reduce_crosstab(counts: pd.DataFrame, top_rows: int = 40, top_cols: int = 40) -> pd.DataFrame:
    c = counts.copy()
    if c.shape[1] > top_cols:
        col_tot = c.sum(axis=0).sort_values(ascending=False)
        keep_cols = col_tot.index[:top_cols].tolist()
        other_cols = [x for x in c.columns if x not in keep_cols]
        c = c[keep_cols].copy()
        if other_cols:
            c["OTROS"] = counts[other_cols].sum(axis=1)

    if c.shape[0] > top_rows:
        row_tot = c.sum(axis=1).sort_values(ascending=False)
        keep_rows = row_tot.index[:top_rows].tolist()
        other_rows = [x for x in c.index if x not in keep_rows]
        c2 = c.loc[keep_rows].copy()
        if other_rows:
            otros = c.loc[other_rows].sum(axis=0).to_frame().T
            otros.index = ["OTROS"]
            c2 = pd.concat([c2, otros], axis=0)
        c = c2

    return c

def apply_bar_labels(fig, show_labels: bool, fmt: str = ".0f"):
    if show_labels:
        fig.update_traces(texttemplate=f"%{{x:{fmt}}}", textposition="outside", cliponaxis=False)
    return fig

def nice_xaxis(fig):
    fig.update_xaxes(type="category", automargin=True)
    fig.update_layout(margin=dict(b=80))
    return fig

def _day_int(ts: pd.Timestamp) -> int:
    return int(np.datetime64(pd.Timestamp(ts).normalize(), "D").astype("int64"))

def _date_from_day_int(x: int) -> pd.Timestamp:
    return pd.Timestamp(np.datetime64(x, "D"))

def _merge_segments(segs: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
    if not segs:
        return []
    segs = sorted(segs, key=lambda x: x[0])
    out = [segs[0]]
    for s, e in segs[1:]:
        ps, pe = out[-1]
        if s <= pe + 1:
            out[-1] = (ps, max(pe, e))
        else:
            out.append((s, e))
    return out

def _intersect_len(seg: Tuple[int, int], allowed: List[Tuple[int, int]]) -> int:
    s, e = seg
    if s > e or not allowed:
        return 0
    tot = 0
    for a, b in allowed:
        ss = max(s, a)
        ee = min(e, b)
        if ss <= ee:
            tot += (ee - ss + 1)
    return tot


# =============================================================================
# Mapping Área y Clasificación
# =============================================================================
def _map_area(area_raw: pd.Series) -> Tuple[pd.Series, pd.Series]:
    key = area_raw.astype("string").str.strip()
    key_u = key.str.upper()

    std = key_u.map(lambda x: AREA_REF.get(x, (None, None))[0] if pd.notna(x) else None)
    gen = key_u.map(lambda x: AREA_REF.get(x, (None, None))[1] if pd.notna(x) else None)

    std = std.fillna(key).replace({"": pd.NA}).fillna(MISSING_LABEL).astype("string")
    gen = gen.fillna(pd.NA).replace({"": pd.NA}).fillna(MISSING_LABEL).astype("string")
    return std, gen

def _map_clas(clas_raw: pd.Series) -> pd.Series:
    key = clas_raw.astype("string").str.strip()
    key_u = key.str.upper()
    std = key_u.map(lambda x: CLAS_REF.get(x, None) if pd.notna(x) else None)
    std = std.fillna(key).replace({"": pd.NA}).fillna(MISSING_LABEL).astype("string")
    return std


# =============================================================================
# Lectura robusta
# =============================================================================
def read_excel_any(file_obj_or_path, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(file_obj_or_path, sheet_name=sheet_name)

def read_excel_strict_hist(file_obj_or_path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(file_obj_or_path, sheet_name=sheet_name)
    cols = list(df.columns)
    r_col = next((c for c in cols if str(c).strip() in R_COL_CANDIDATES), None)
    keep = [c for c in REQUIRED_COLS if c in cols]
    if r_col and r_col not in keep:
        keep.append(r_col)
    return df[keep].copy() if keep else df.copy()

def read_csv_any(file_obj_or_path) -> pd.DataFrame:
    return pd.read_csv(file_obj_or_path)

def find_cost_sheet_name(xls: pd.ExcelFile) -> Optional[str]:
    for s in xls.sheet_names:
        if str(s).strip().lower() == "costo nominal":
            return s
    for s in xls.sheet_names:
        ss = str(s).strip().lower()
        if "costo" in ss and "nominal" in ss:
            return s
    return None


# =============================================================================
# Preparación de Historia Personal
# =============================================================================
KEEP_INTERNAL = [
    "cod", "ini", "fin", "fin_eff", "fnac",
    "r_pct",
    "clas_raw", "clas",
    "sexo", "ts", "emp",
    "area_raw", "area", "area_gen",
    "cargo", "nac", "lug", "reg",
]

@st.cache_data(show_spinner=False)
def validate_and_prepare_hist(df_raw: pd.DataFrame) -> pd.DataFrame:
    missing = [c for c in REQUIRED_COLS if c not in df_raw.columns]
    if missing:
        raise ValueError("Faltan columnas requeridas en Historia Personal:\n- " + "\n- ".join(missing))

    cols = list(df_raw.columns)
    r_col = next((c for c in cols if str(c).strip() in R_COL_CANDIDATES), None)

    use_cols = REQUIRED_COLS.copy()
    if r_col and r_col not in use_cols:
        use_cols.append(r_col)

    df = df_raw[use_cols].copy()
    out = df.rename(columns=COL_MAP)

    out["ini"] = _to_datetime(out["ini"])
    out["fin"] = _to_datetime(out["fin"])
    out["fnac"] = _to_datetime(out["fnac"])
    out["fin_eff"] = out["fin"].fillna(today_dt())

    if r_col:
        out = out.rename(columns={r_col: "r_pct"})
    else:
        out["r_pct"] = 1.0

    for c in ["cod", "clas_raw", "sexo", "ts", "emp", "area_raw", "cargo", "nac", "lug", "reg"]:
        out[c] = out[c].astype("string").str.strip()
        out.loc[out[c].isin(["", "None", "nan", "NaT"]), c] = pd.NA

    out = out[~out["cod"].isna()].copy()
    out = out[~out["ini"].isna()].copy()
    out["cod"] = out["cod"].astype(str)

    rp = out["r_pct"].copy()
    if rp.dtype == "object" or str(rp.dtype).startswith("string"):
        rp2 = rp.astype(str).str.replace("%", "", regex=False).str.strip()
        rp_num = pd.to_numeric(rp2, errors="coerce")
        rp_num = np.where(rp_num > 1.5, rp_num / 100.0, rp_num)
        out["r_pct"] = pd.Series(rp_num, index=out.index).fillna(1.0).astype(float)
    else:
        rp_num = pd.to_numeric(rp, errors="coerce").fillna(1.0).astype(float)
        rp_num = np.where(rp_num > 1.5, rp_num / 100.0, rp_num)
        out["r_pct"] = rp_num

    out["area"], out["area_gen"] = _map_area(out["area_raw"])
    out["clas"] = _map_clas(out["clas_raw"])

    out = out[KEEP_INTERNAL].copy()
    out = out.sort_values(["cod", "ini", "fin_eff"]).reset_index(drop=True)
    return out


# =============================================================================
# Preparación Costo Nominal
# =============================================================================
@st.cache_data(show_spinner=False)
def validate_and_prepare_cost(df_cost_raw: pd.DataFrame) -> pd.DataFrame:
    missing = [c for c in COST_REQUIRED if c not in df_cost_raw.columns]
    if missing:
        raise ValueError("Faltan columnas requeridas en Costo Nominal:\n- " + "\n- ".join(missing))

    d = df_cost_raw[COST_REQUIRED].copy().rename(columns=COST_COL_MAP)
    d["cod"] = d["cod"].astype("string").str.strip().astype(str)
    d["c_ini"] = _to_datetime(d["c_ini"])
    d["c_fin"] = _to_datetime(d["c_fin"]).fillna(today_dt())
    d["costo"] = pd.to_numeric(d["costo"], errors="coerce").fillna(0.0).astype(float)

    d = d[~d["cod"].isna()].copy()
    d = d[~d["c_ini"].isna()].copy()
    d = d.sort_values(["cod", "c_ini", "c_fin"]).reset_index(drop=True)
    d = d[d["c_fin"] >= d["c_ini"]].copy()
    return d


# =============================================================================
# Intervalos por persona (para existencias globales legacy PDE)
# =============================================================================
@st.cache_data(show_spinner=False)
def merge_intervals_per_person(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for cod, g in df.groupby("cod", sort=False):
        g = g.sort_values(["ini", "fin_eff"]).copy()

        cur_ini = None
        cur_fin = None
        cur_row = None

        for _, r in g.iterrows():
            ini = r["ini"]
            fin = r["fin_eff"]

            if cur_ini is None:
                cur_ini, cur_fin = ini, fin
                cur_row = r
                continue

            if ini <= (cur_fin + pd.Timedelta(days=1)):
                if fin > cur_fin:
                    cur_fin = fin
                cur_row = r
            else:
                out_r = cur_row.copy()
                out_r["ini"] = cur_ini
                out_r["fin_eff"] = cur_fin
                rows.append(out_r)

                cur_ini, cur_fin = ini, fin
                cur_row = r

        if cur_ini is not None:
            out_r = cur_row.copy()
            out_r["ini"] = cur_ini
            out_r["fin_eff"] = cur_fin
            rows.append(out_r)

    return pd.DataFrame(rows).reset_index(drop=True)


# =============================================================================
# Buckets (tenure/age)
# =============================================================================
TENURE_BUCKETS = {
    "< 30 días": (0, 29),
    "30 - 90 días": (30, 90),
    "91 - 180 días": (91, 180),
    "181 - 360 días": (181, 360),
    "> 360 días": (361, None),
}

AGE_BUCKETS = {
    "< 24 años": (None, 23),
    "24 - 30 años": (24, 30),
    "31 - 37 años": (31, 37),
    "38 - 42 años": (38, 42),
    "43 - 49 años": (43, 49),
    "50 - 56 años": (50, 56),
    "> 56 años": (57, None),
}

def bucket_antiguedad(days: pd.Series) -> pd.Series:
    d = days.astype("float")
    out = pd.Series(np.where(d.isna(), MISSING_LABEL, ""), index=days.index, dtype="object")
    out = np.where((~pd.isna(d)) & (d >= 0) & (d < 30), "< 30 días", out)
    out = np.where((~pd.isna(d)) & (d >= 30) & (d <= 90), "30 - 90 días", out)
    out = np.where((~pd.isna(d)) & (d >= 91) & (d <= 180), "91 - 180 días", out)
    out = np.where((~pd.isna(d)) & (d >= 181) & (d <= 360), "181 - 360 días", out)
    out = np.where((~pd.isna(d)) & (d >= 361), "> 360 días", out)
    return pd.Series(out, index=days.index, dtype="object")

def bucket_edad_from_dob(dob: pd.Series, ref: pd.Series) -> pd.Series:
    out = pd.Series(MISSING_LABEL, index=dob.index, dtype="object")
    mask = (~dob.isna()) & (~ref.isna())
    if not mask.any():
        return out

    dob2 = dob[mask]
    ref2 = ref[mask]

    had_bday = (ref2.dt.month > dob2.dt.month) | ((ref2.dt.month == dob2.dt.month) & (ref2.dt.day >= dob2.dt.day))
    edad = (ref2.dt.year - dob2.dt.year) - (~had_bday).astype(int)

    out.loc[mask] = np.where(edad < 24, "< 24 años", out.loc[mask])
    out.loc[mask] = np.where((edad >= 24) & (edad <= 30), "24 - 30 años", out.loc[mask])
    out.loc[mask] = np.where((edad >= 31) & (edad <= 37), "31 - 37 años", out.loc[mask])
    out.loc[mask] = np.where((edad >= 38) & (edad <= 42), "38 - 42 años", out.loc[mask])
    out.loc[mask] = np.where((edad >= 43) & (edad <= 49), "43 - 49 años", out.loc[mask])
    out.loc[mask] = np.where((edad >= 50) & (edad <= 56), "50 - 56 años", out.loc[mask])
    out.loc[mask] = np.where((edad >= 57), "> 56 años", out.loc[mask])
    return out


# =============================================================================
# Filtros (multi-select)
# =============================================================================
@dataclass
class FilterState:
    sexo: List[str]
    area_gen: List[str]
    area: List[str]
    cargo: List[str]
    clas: List[str]
    ts: List[str]
    emp: List[str]
    nac: List[str]
    lug: List[str]
    reg: List[str]
    antig: List[str]  # tenure buckets seleccionados (filtra persona-días y eventos a esos bins)
    edad: List[str]   # age buckets

def apply_categorical_filters(df: pd.DataFrame, fs: FilterState) -> pd.DataFrame:
    out = df.copy()

    def _apply(col: str, selected: List[str]) -> None:
        nonlocal out
        if selected:
            out = out[out[col].isin(selected)]

    _apply("sexo", fs.sexo)
    _apply("area_gen", fs.area_gen)
    _apply("area", fs.area)
    _apply("cargo", fs.cargo)
    _apply("clas", fs.clas)
    _apply("ts", fs.ts)
    _apply("emp", fs.emp)
    _apply("nac", fs.nac)
    _apply("lug", fs.lug)
    _apply("reg", fs.reg)

    return out


# =============================================================================
# Period windows (D/W/M/Y)
# =============================================================================
def build_period_windows(start: pd.Timestamp, end: pd.Timestamp, period: str) -> pd.DataFrame:
    days = pd.date_range(start, end, freq="D")
    cal = add_calendar_fields(pd.DataFrame({"Día": days}), "Día")

    if period == "D":
        w = cal[["Día"]].copy()
        w["window_start"] = w["Día"]
        w["window_end"] = w["Día"]
        w["cut"] = w["Día"]
        w["Periodo"] = w["Día"].dt.strftime("%Y-%m-%d")
        return w[["Periodo", "cut", "window_start", "window_end"]]

    if period == "W":
        w = cal.groupby("CodSem", as_index=False).agg(
            window_start=("Día", "min"),
            window_end=("Día", "max"),
            cut=("FinSemana", "max"),
        )
        w["Periodo"] = w["CodSem"].astype(str)
        return w[["Periodo", "cut", "window_start", "window_end"]].sort_values("cut")

    if period == "M":
        w = cal.groupby("CodMes", as_index=False).agg(
            window_start=("Día", "min"),
            window_end=("Día", "max"),
            cut=("FinMes", "max"),
        )
        w["Periodo"] = w["CodMes"].astype(str)
        return w[["Periodo", "cut", "window_start", "window_end"]].sort_values("cut")

    if period == "Y":
        w = cal.groupby("Año", as_index=False).agg(
            window_start=("Día", "min"),
            window_end=("Día", "max"),
            cut=("Día", "max"),
        )
        w["Periodo"] = w["Año"].astype(int).astype(str)
        return w[["Periodo", "cut", "window_start", "window_end"]].sort_values("cut")

    raise ValueError("period inválido")

def period_label_series(d: pd.Series, period: str) -> pd.Series:
    dd = pd.to_datetime(d, errors="coerce").dt.normalize()
    if period == "D":
        return dd.dt.strftime("%Y-%m-%d")
    if period == "W":
        yr = dd.dt.year.astype(int)
        wk = excel_weeknum_return_type_1(dd).astype(int)
        yy = (yr % 100).astype(int).astype(str).str.zfill(2)
        ww = wk.astype(int).astype(str).str.zfill(2)
        return (yy + ww).astype(str)
    if period == "M":
        yr = dd.dt.year.astype(int)
        mo = dd.dt.month.astype(int)
        yy = (yr % 100).astype(int).astype(str).str.zfill(2)
        mm = mo.astype(int).astype(str).str.zfill(2)
        return (yy + mm).astype(str)
    if period == "Y":
        return dd.dt.year.astype("Int64").astype(str)
    raise ValueError("period inválido")


# =============================================================================
# 1) MÉTRICA NUEVA: ROTACIÓN ROBUSTA POR ANTIGÜEDAD (persona-día)
# =============================================================================
def _age_allowed_segments(dob_ts, start_day: int, end_day: int, edad_sel: List[str]) -> List[Tuple[int, int]]:
    if not edad_sel:
        return [(start_day, end_day)]

    if pd.isna(dob_ts):
        return [(start_day, end_day)] if (MISSING_LABEL in edad_sel) else []

    dob_date = pd.Timestamp(dob_ts).date()
    segs: List[Tuple[int, int]] = []
    for b in edad_sel:
        if b == MISSING_LABEL:
            continue
        if b not in AGE_BUCKETS:
            continue
        y0, y1 = AGE_BUCKETS[b]
        s_date = date(1900, 1, 1) if y0 is None else _years_offset_date(dob_date, y0)
        e_date = date(3000, 1, 1) if y1 is None else (_years_offset_date(dob_date, y1 + 1) - timedelta(days=1))
        s = max(_day_int(pd.Timestamp(s_date)), start_day)
        e = min(_day_int(pd.Timestamp(e_date)), end_day)
        if s <= e:
            segs.append((s, e))
    return _merge_segments(segs)

def _tenure_bin_segments_for_spell(
    ini_day: int,
    base_seg: Tuple[int, int],
    tenure_bins: List[Tuple[str, int, Optional[int]]],
) -> Dict[str, Tuple[int, int]]:
    # devuelve seg por bin en coordenadas absolutas (día int) intersectado con base_seg
    bs, be = base_seg
    out: Dict[str, Tuple[int, int]] = {}
    for lab, a0, a1 in tenure_bins:
        s = ini_day + int(a0)
        e = be if a1 is None else (ini_day + int(a1))
        ss = max(bs, s)
        ee = min(be, e)
        if ss <= ee:
            out[lab] = (ss, ee)
    return out

def _assign_tenure_bin(tenure_days: int, tenure_bins: List[Tuple[str, int, Optional[int]]]) -> Optional[str]:
    for lab, a0, a1 in tenure_bins:
        if tenure_days < a0:
            continue
        if a1 is None:
            return lab
        if a0 <= tenure_days <= a1:
            return lab
    return None

@st.cache_data(show_spinner=False)
def compute_rr_period_table(
    df_events: pd.DataFrame,
    start: pd.Timestamp,
    end: pd.Timestamp,
    period: str,
    tenure_bins: List[Tuple[str, int, Optional[int]]],
    antig_sel: List[str],  # si no vacío, limita a estos bins
    edad_sel: List[str],   # filtra persona-días (exposure) y eventos (exits) por edad
    unique_salidas_por_dia: bool,
) -> pd.DataFrame:
    """
    Retorna tabla por periodo x bin:
      Exposure (persona-días), Exits (conteo), Exits_w (sum r_pct),
      Rate_1000 = Exits/Exposure*1000
      RateW_1000 = Exits_w/Exposure*1000
    + agrega filas Bin='TOTAL' por periodo (sum bins).
    """
    if df_events.empty:
        return pd.DataFrame()

    start = pd.Timestamp(start).normalize()
    end = pd.Timestamp(end).normalize()
    if start > end:
        return pd.DataFrame()

    # bins elegidos para el cálculo (si usuario selecciona antig buckets)
    allowed_bins = [b for b, _, _ in tenure_bins]
    if antig_sel:
        allowed_bins = [b for b in allowed_bins if b in set(antig_sel)]
        if not allowed_bins:
            return pd.DataFrame()

    tbins = [t for t in tenure_bins if t[0] in allowed_bins]

    windows = build_period_windows(start, end, period).copy()
    if windows.empty:
        return pd.DataFrame()

    windows["ws_day"] = windows["window_start"].apply(_day_int)
    windows["we_day"] = windows["window_end"].apply(_day_int)

    p_labels = windows["Periodo"].tolist()
    p_ws = windows["ws_day"].to_numpy(np.int64)
    p_we = windows["we_day"].to_numpy(np.int64)

    # Exposure acumuladores
    expo = {(p, b): 0.0 for p in p_labels for b in allowed_bins}

    dfe = df_events[(df_events["ini"] <= end) & (df_events["fin_eff"] >= start)].copy()
    if dfe.empty:
        # aún puede haber exits sin exposure? igual devolvemos vacío
        return pd.DataFrame()

    dfe["ini_day"] = dfe["ini"].apply(_day_int).astype(np.int64)
    dfe["fin_day"] = dfe["fin_eff"].apply(_day_int).astype(np.int64)

    start_day = _day_int(start)
    end_day = _day_int(end)

    # Loop spells -> acumula exposure por periodo y bin (sin expandir día)
    for r in dfe.itertuples(index=False):
        ini_day = int(r.ini_day)
        fin_day = int(r.fin_day)

        base_s = max(ini_day, start_day)
        base_e = min(fin_day, end_day)
        if base_s > base_e:
            continue

        age_allowed = _age_allowed_segments(r.fnac, start_day, end_day, edad_sel)
        if edad_sel and not age_allowed:
            continue

        overlap_idx = np.where((p_ws <= base_e) & (p_we >= base_s))[0]
        if overlap_idx.size == 0:
            continue

        for j in overlap_idx:
            ws = int(max(p_ws[j], base_s))
            we = int(min(p_we[j], base_e))
            if ws > we:
                continue
            base_seg = (ws, we)
            segs_by_bin = _tenure_bin_segments_for_spell(ini_day, base_seg, tbins)
            for b, seg in segs_by_bin.items():
                ln = _intersect_len(seg, age_allowed) if edad_sel else (seg[1] - seg[0] + 1)
                if ln > 0:
                    expo[(p_labels[j], b)] += float(ln)

    # Exits: numerador por periodo x bin (asignado a bin por antigüedad al salir)
    exits = df_events[~df_events["fin"].isna()].copy()
    exits = exits[(exits["fin"] >= start) & (exits["fin"] <= end)].copy()
    if not exits.empty:
        exits["exit_day"] = exits["fin"].apply(_day_int).astype(np.int64)

        exits["tenure_days_exit"] = (exits["fin"] - exits["ini"]).dt.days.astype("Int64")
        exits = exits[~exits["tenure_days_exit"].isna()].copy()
        exits["tenure_days_exit"] = exits["tenure_days_exit"].astype(int)

        # edad al salir
        exits["Edad"] = bucket_edad_from_dob(exits["fnac"], exits["fin"])
        if edad_sel:
            if MISSING_LABEL in edad_sel:
                exits = exits[exits["Edad"].isin(edad_sel)]
            else:
                exits = exits[(exits["Edad"].isin(edad_sel)) & (exits["Edad"] != MISSING_LABEL)]

        exits["Bin"] = exits["tenure_days_exit"].apply(lambda x: _assign_tenure_bin(int(x), tbins))
        exits = exits[~exits["Bin"].isna()].copy()
        exits = exits[exits["Bin"].isin(allowed_bins)].copy()

        exits["Periodo"] = period_label_series(exits["fin"], period)

        # unicidad por día (opción legacy) — aquí solo por seguridad
        if unique_salidas_por_dia:
            exits = exits.sort_values(["cod", "exit_day"]).drop_duplicates(["cod", "exit_day"], keep="last")

        agg_ex = exits.groupby(["Periodo", "Bin"], as_index=False).agg(
            Exits=("cod", "size"),
            Exits_w=("r_pct", "sum"),
        )
    else:
        agg_ex = pd.DataFrame(columns=["Periodo", "Bin", "Exits", "Exits_w"])

    # Construir tabla completa period x bin
    rows = []
    for p in p_labels:
        for b in allowed_bins:
            rows.append({"Periodo": p, "Bin": b, "Exposure": expo[(p, b)]})
    out = pd.DataFrame(rows)

    out = out.merge(agg_ex, on=["Periodo", "Bin"], how="left")
    out["Exits"] = out["Exits"].fillna(0).astype(float)
    out["Exits_w"] = out["Exits_w"].fillna(0.0).astype(float)

    out["Rate_1000"] = np.where(out["Exposure"] > 0, (out["Exits"] / out["Exposure"]) * 1000.0, np.nan)
    out["RateW_1000"] = np.where(out["Exposure"] > 0, (out["Exits_w"] / out["Exposure"]) * 1000.0, np.nan)

    # total por periodo (suma bins)
    tot = out.groupby("Periodo", as_index=False).agg(
        Exposure=("Exposure", "sum"),
        Exits=("Exits", "sum"),
        Exits_w=("Exits_w", "sum"),
    )
    tot["Bin"] = "TOTAL"
    tot["Rate_1000"] = np.where(tot["Exposure"] > 0, (tot["Exits"] / tot["Exposure"]) * 1000.0, np.nan)
    tot["RateW_1000"] = np.where(tot["Exposure"] > 0, (tot["Exits_w"] / tot["Exposure"]) * 1000.0, np.nan)

    out = pd.concat([out, tot], ignore_index=True)

    # anexar orden (cut) desde windows
    out = out.merge(windows[["Periodo", "cut", "window_start", "window_end"]], on="Periodo", how="left")
    out = out.sort_values(["cut", "Bin"]).reset_index(drop=True)
    return out

@st.cache_data(show_spinner=False)
def compute_rr_kpi(
    rr_group: pd.DataFrame,
    rr_base: pd.DataFrame,
    period: str,
    green_max: float,
    yellow_max: float,
) -> pd.DataFrame:
    """
    Meta estacional por periodo y bin:
      Meta = avg(shift(S), shift(S+1), shift(S+2)) sobre baseline RateW_1000
    KPI_RR = RateW_1000 / Meta
    Semáforo: verde si ≤ green_max, amarillo si ≤ yellow_max, rojo si > yellow_max
    """
    if rr_group.empty:
        return pd.DataFrame()

    if period == "W":
        S = 52
    elif period == "M":
        S = 12
    elif period == "D":
        S = 365
    else:
        S = 1

    b = rr_base.copy()
    g = rr_group.copy()

    b = b.sort_values(["Bin", "cut"]).reset_index(drop=True)
    b["Meta"] = b.groupby("Bin")["RateW_1000"].transform(lambda x: (x.shift(S) + x.shift(S + 1) + x.shift(S + 2)) / 3.0)

    m = g.merge(b[["Periodo", "Bin", "Meta"]], on=["Periodo", "Bin"], how="left")

    eps = 1e-12
    m["KPI_RR"] = m["RateW_1000"] / (m["Meta"] + eps)

    def _semaforo(k: float) -> str:
        if np.isnan(k):
            return "-"
        if k <= green_max:
            return "VERDE"
        if k <= yellow_max:
            return "AMARILLO"
        return "ROJO"

    m["Semaforo"] = m["KPI_RR"].apply(_semaforo)
    return m.sort_values(["cut", "Bin"]).reset_index(drop=True)


# =============================================================================
# 2) Supervivencia (Kaplan–Meier simple)
# =============================================================================
@st.cache_data(show_spinner=False)
def km_survival_curve(
    df_events: pd.DataFrame,
    t0: pd.Timestamp,
    t1: pd.Timestamp,
    H_days: int,
) -> Tuple[pd.DataFrame, float]:
    """
    KM simple por cohortes: spells con ini en [t0,t1].
    Tiempo = min(fin, t1) - ini (días)
    Evento observado si fin no-null y fin <= t1
    Retorna curva (day, S) hasta H_days y S(H_days).
    """
    t0 = pd.Timestamp(t0).normalize()
    t1 = pd.Timestamp(t1).normalize()
    if df_events.empty:
        return pd.DataFrame({"day": [], "S": []}), np.nan

    cohort = df_events[(df_events["ini"] >= t0) & (df_events["ini"] <= t1)].copy()
    if cohort.empty:
        return pd.DataFrame({"day": [], "S": []}), np.nan

    # tiempo y evento
    fin_obs = cohort["fin"].copy()
    cens_end = t1
    end_time = fin_obs.fillna(cens_end).clip(upper=cens_end)
    T = (end_time - cohort["ini"]).dt.days.astype(int)
    E = (~fin_obs.isna()) & (fin_obs <= t1)

    T = T.clip(lower=0)
    df = pd.DataFrame({"T": T, "E": E.astype(int)})

    # KM
    # n_at_risk(t) = # con T >= t
    # d(t) = # eventos en t
    df_ev = df[df["E"] == 1].copy()
    if df_ev.empty:
        curve = pd.DataFrame({"day": [0, H_days], "S": [1.0, 1.0]})
        return curve, 1.0

    times = np.sort(df_ev["T"].unique())
    S = 1.0
    rows = [{"day": 0, "S": 1.0}]
    for t in times:
        n = (df["T"] >= t).sum()
        d = ((df["T"] == t) & (df["E"] == 1)).sum()
        if n <= 0:
            continue
        S *= (1.0 - (d / n))
        rows.append({"day": int(t), "S": float(S)})

    curve = pd.DataFrame(rows).sort_values("day").reset_index(drop=True)

    # recorte a H
    curve = curve[curve["day"] <= H_days].copy()
    if curve.empty:
        curve = pd.DataFrame({"day": [0], "S": [1.0]})

    # S(H): último valor <= H
    S_H = float(curve.iloc[-1]["S"]) if not curve.empty else np.nan

    # agrega punto final en H si no existe
    if int(curve.iloc[-1]["day"]) < H_days:
        curve = pd.concat([curve, pd.DataFrame([{"day": H_days, "S": S_H}])], ignore_index=True)

    return curve, S_H


# =============================================================================
# LEGACY: PDE y KPI_COST (se mantiene tal cual)
# =============================================================================
def compute_existencias_daily_filtered_fast(
    df_intervals: pd.DataFrame,
    start: pd.Timestamp,
    end: pd.Timestamp,
    antig_sel: List[str],
    edad_sel: List[str],
) -> pd.DataFrame:
    idx = pd.date_range(start, end, freq="D")
    n = len(idx)
    if n == 0:
        return pd.DataFrame({"Día": [], "Existencias": []})

    g = df_intervals[(df_intervals["ini"] <= end) & (df_intervals["fin_eff"] >= start)].copy()
    if g.empty:
        out = pd.DataFrame({"Día": idx, "Existencias": np.zeros(n, dtype=int)})
        return add_calendar_fields(out, "Día")

    use_antig = bool(antig_sel)
    use_edad = bool(edad_sel)

    ini_days = g["ini"].values.astype("datetime64[D]").astype("int64")
    fin_days = g["fin_eff"].values.astype("datetime64[D]").astype("int64")
    start_day = np.datetime64(start, "D").astype("int64")
    end_day = np.datetime64(end, "D").astype("int64")

    diff = np.zeros(n + 1, dtype=np.int64)

    antig_list = [b for b in antig_sel if b in TENURE_BUCKETS] if use_antig else []
    edad_list = [b for b in edad_sel if b in AGE_BUCKETS] if use_edad else []
    edad_allow_sindato = use_edad and (MISSING_LABEL in edad_sel)

    fnac_vals = g["fnac"].values

    for i in range(len(g)):
        base_s = max(ini_days[i], start_day)
        base_e = min(fin_days[i], end_day)
        if base_s > base_e:
            continue

        if (not use_antig) and (not use_edad):
            s_idx = int(base_s - start_day)
            e_idx = int(base_e - start_day)
            diff[s_idx] += 1
            diff[e_idx + 1 if (e_idx + 1 < n) else n] -= 1
            continue

        dob_ts = fnac_vals[i]
        dob_missing = pd.isna(dob_ts)
        if use_edad and dob_missing and not edad_allow_sindato:
            continue

        ini0 = ini_days[i]

        # Antig ranges
        if use_antig and antig_list:
            antig_ranges = []
            for b in antig_list:
                a0, a1 = TENURE_BUCKETS[b]
                s = max(ini0 + a0, base_s)
                e = min(base_e, (base_e if a1 is None else ini0 + a1))
                if s <= e:
                    antig_ranges.append((s, e))
            if not antig_ranges:
                continue
        else:
            antig_ranges = [(base_s, base_e)]

        # Edad ranges
        if use_edad and (not dob_missing) and edad_list:
            dob_date = pd.Timestamp(dob_ts).date()
            edad_ranges = []
            for b in edad_list:
                y0, y1 = AGE_BUCKETS[b]
                s_date = start.date() if y0 is None else _years_offset_date(dob_date, y0)
                e_date = end.date() if y1 is None else (_years_offset_date(dob_date, y1 + 1) - timedelta(days=1))

                s = max(np.int64(np.datetime64(s_date, "D").astype("int64")), base_s)
                e = min(np.int64(np.datetime64(e_date, "D").astype("int64")), base_e)
                if s <= e:
                    edad_ranges.append((s, e))
            if not edad_ranges:
                continue
        else:
            edad_ranges = [(base_s, base_e)]

        # Intersección -> diff
        for (as_, ae_) in antig_ranges:
            for (es_, ee_) in edad_ranges:
                s = max(as_, es_)
                e = min(ae_, ee_)
                if s <= e:
                    s_idx = int(s - start_day)
                    e_idx = int(e - start_day)
                    diff[s_idx] += 1
                    diff[e_idx + 1 if (e_idx + 1 < n) else n] -= 1

    exist = np.cumsum(diff[:-1]).astype(int)
    out = pd.DataFrame({"Día": idx, "Existencias": exist})
    return add_calendar_fields(out, "Día")

def compute_salidas_daily_filtered(
    df_events: pd.DataFrame,
    start: pd.Timestamp,
    end: pd.Timestamp,
    antig_sel: List[str],
    edad_sel: List[str],
    unique_personas_por_dia: bool = True,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    idx = pd.date_range(start, end, freq="D")
    if len(idx) == 0:
        return pd.DataFrame({"Día": [], "Salidas": []}), df_events.iloc[0:0].copy()

    d = df_events[~df_events["fin"].isna()].copy()
    d = d[(d["fin"] >= start) & (d["fin"] <= end)].copy()

    if d.empty:
        out = pd.DataFrame({"Día": idx, "Salidas": np.zeros(len(idx), dtype=int)})
        return add_calendar_fields(out, "Día"), d

    d = d.rename(columns={"fin": "ref_fin"})
    d["ref_fin"] = _to_datetime(d["ref_fin"])

    d["antig_dias"] = (d["ref_fin"] - d["ini"]).dt.days
    d["Antigüedad"] = bucket_antiguedad(d["antig_dias"])
    d["Edad"] = bucket_edad_from_dob(d["fnac"], d["ref_fin"])

    if antig_sel:
        d = d[d["Antigüedad"].isin(antig_sel)]
    if edad_sel:
        d = d[d["Edad"].isin(edad_sel)]

    if unique_personas_por_dia:
        g = d.groupby("ref_fin")["cod"].nunique().rename("Salidas")
    else:
        g = d.groupby("ref_fin")["cod"].size().rename("Salidas")

    out = pd.DataFrame({"Día": idx}).merge(
        g.reset_index().rename(columns={"ref_fin": "Día"}),
        on="Día",
        how="left",
    )
    out["Salidas"] = out["Salidas"].fillna(0).astype(int)
    out = add_calendar_fields(out, "Día")
    return out, d

def aggregate_daily_to_period_for_pde(df_daily_g: pd.DataFrame, period: str) -> pd.DataFrame:
    d = df_daily_g.copy()
    if "CodSem" not in d.columns or "CodMes" not in d.columns or "Año" not in d.columns:
        d = add_calendar_fields(d, "Día")

    key = {"D": "Día", "W": "CodSem", "M": "CodMes", "Y": "Año"}[period]
    cut_col = {"D": "Día", "W": "FinSemana", "M": "FinMes", "Y": "Día"}[period]

    if "Salidas" not in d.columns:
        d["Salidas"] = 0
    if "Existencias" not in d.columns:
        d["Existencias"] = 0

    def _agg_group(g: pd.DataFrame) -> pd.Series:
        ws = g["Día"].min()
        we = g["Día"].max()
        L = int((we - ws).days + 1) if pd.notna(ws) and pd.notna(we) else 0

        expo = float(np.nansum(g["Existencias"].astype(float).values))
        sal = float(np.nansum(g["Salidas"].astype(float).values))

        delta = (we - g["Día"]).dt.days.astype(float)
        perd = float(np.nansum(g["Salidas"].astype(float).values * delta.values))

        pot = expo + perd
        pde = (perd / pot) if pot > 0 else np.nan
        exist_prom = (expo / L) if L > 0 else np.nan

        cut = g[cut_col].max() if cut_col in g.columns else we
        return pd.Series({
            "window_start": ws,
            "window_end": we,
            "cut": cut,
            "Dias_Periodo": L,
            "Salidas": sal,
            "Exposicion": expo,
            "Perdidos": perd,
            "Potencial": pot,
            "PDE": pde,
            "Existencias_Prom": exist_prom,
        })

    agg = d.groupby(key, dropna=False, as_index=False).apply(_agg_group).reset_index(drop=True)

    if period == "D":
        agg["Periodo"] = pd.to_datetime(agg["window_start"]).dt.strftime("%Y-%m-%d")
    elif period in ("W", "M"):
        agg["Periodo"] = agg[key].astype(str)
    else:
        agg["Periodo"] = agg[key].astype(int).astype(str)

    agg = agg.sort_values("cut").reset_index(drop=True)
    return agg

@st.cache_data(show_spinner=False)
def compute_pde_kpis(
    df_period_g: pd.DataFrame,
    df_period_base: pd.DataFrame,
    horizon_days: int,
    period: str,
    green_max: float = 0.95,
    yellow_max: float = 1.05,
) -> pd.DataFrame:
    eps = 1e-12
    H = float(horizon_days)

    g = df_period_g.copy()
    b = df_period_base.copy()

    keep_cols = [
        "Periodo", "cut", "window_start", "window_end", "Dias_Periodo",
        "Salidas", "Exposicion", "Perdidos", "Potencial", "PDE", "Existencias_Prom"
    ]
    for df_ in (g, b):
        if "Dias_Periodo" not in df_.columns:
            df_["Dias_Periodo"] = (pd.to_datetime(df_["window_end"]) - pd.to_datetime(df_["window_start"])).dt.days + 1

    g = g[[c for c in keep_cols if c in g.columns]].rename(columns={
        "Salidas": "Salidas_g",
        "Exposicion": "Obs_g",
        "Perdidos": "Perdidos_g",
        "Potencial": "Pot_g",
        "PDE": "PDE_g",
        "Existencias_Prom": "HC_g_prom",
    })
    b = b[[c for c in keep_cols if c in b.columns]].rename(columns={
        "Salidas": "Salidas_b",
        "Exposicion": "Obs_b",
        "Perdidos": "Perdidos_b",
        "Potencial": "Pot_b",
        "PDE": "PDE_b",
        "Existencias_Prom": "HC_b_prom",
    })

    m = g.merge(b, on=["Periodo", "cut", "window_start", "window_end", "Dias_Periodo"], how="left")

    L = m["Dias_Periodo"].astype(float).replace(0, np.nan)
    m["PDEH_g"] = 1.0 - np.power((1.0 - m["PDE_g"].astype(float)).clip(lower=0.0, upper=1.0), (H / (L + eps)))
    m["PDEH_b"] = 1.0 - np.power((1.0 - m["PDE_b"].astype(float)).clip(lower=0.0, upper=1.0), (H / (L + eps)))

    if period == "W":
        S = 52
    elif period == "M":
        S = 12
    elif period == "D":
        S = 365
    else:
        S = 1

    p = m["PDEH_b"].astype(float)
    m["Meta"] = (p.shift(S) + p.shift(S + 1) + p.shift(S + 2)) / 3.0

    m["KPI_PDE"] = m["PDEH_g"] / (m["Meta"] + eps)
    m["Brecha_vs_Meta"] = m["PDEH_g"] - m["Meta"]
    m["Mejora"] = m["KPI_PDE"].shift(1) - m["KPI_PDE"]

    def _semaforo(k: float) -> str:
        if np.isnan(k):
            return "-"
        if k <= green_max:
            return "VERDE"
        if k <= yellow_max:
            return "AMARILLO"
        return "ROJO"

    m["Semaforo"] = m["KPI_PDE"].apply(_semaforo)
    return m.sort_values("cut").reset_index(drop=True)

@st.cache_data(show_spinner=False)
def compute_cost_period_metrics(
    df_events: pd.DataFrame,
    df_cost: pd.DataFrame,
    start: pd.Timestamp,
    end: pd.Timestamp,
    period: str,
    antig_sel: List[str],
    edad_sel: List[str],
    unique_salidas_por_dia: bool,
) -> pd.DataFrame:
    if df_cost is None or df_cost.empty:
        return pd.DataFrame()

    windows = build_period_windows(start, end, period).copy()
    if windows.empty:
        return pd.DataFrame()

    c = df_cost.copy()
    c["ini_day"] = c["c_ini"].apply(_day_int)
    c["fin_day"] = c["c_fin"].apply(_day_int)

    cost_map: Dict[str, Tuple[np.ndarray, np.ndarray, np.ndarray]] = {}
    for cod, g in c.groupby("cod", sort=False):
        cost_map[str(cod)] = (
            g["ini_day"].to_numpy(dtype=np.int64),
            g["fin_day"].to_numpy(dtype=np.int64),
            g["costo"].to_numpy(dtype=float),
        )

    spells = df_events[(df_events["ini"] <= end) & (df_events["fin_eff"] >= start)].copy()
    if spells.empty:
        out = windows.copy()
        out["WorkedCost"] = 0.0
        out["LostCost"] = 0.0
        out["PotentialCost"] = 0.0
        out["KPI_COST"] = np.nan
        out["LostRate"] = np.nan
        return out

    spells["ini_day"] = spells["ini"].apply(_day_int)
    spells["fin_day"] = spells["fin_eff"].apply(_day_int)
    spells["r_pct"] = pd.to_numeric(spells["r_pct"], errors="coerce").fillna(1.0).astype(float)

    exits = df_events[~df_events["fin"].isna()].copy()
    exits = exits[(exits["fin"] >= start) & (exits["fin"] <= end)].copy()
    if not exits.empty:
        exits["ref_fin"] = _to_datetime(exits["fin"])
        exits["antig_dias"] = (exits["ref_fin"] - exits["ini"]).dt.days
        exits["Antigüedad"] = bucket_antiguedad(exits["antig_dias"])
        exits["Edad"] = bucket_edad_from_dob(exits["fnac"], exits["ref_fin"])
        if antig_sel:
            exits = exits[exits["Antigüedad"].isin(antig_sel)]
        if edad_sel:
            exits = exits[exits["Edad"].isin(edad_sel)]
        exits["exit_day"] = exits["ref_fin"].apply(_day_int)
        exits["r_pct"] = pd.to_numeric(exits["r_pct"], errors="coerce").fillna(1.0).astype(float)
        if unique_salidas_por_dia:
            exits = exits.sort_values(["cod", "exit_day"]).drop_duplicates(["cod", "exit_day"], keep="last")
    else:
        exits = exits.iloc[0:0].copy()

    def _segments_for_spell_with_buckets(
        ini_day: int,
        base_s: int,
        base_e: int,
        dob_ts,
        antig_sel: List[str],
        edad_sel: List[str],
        start: pd.Timestamp,
        end: pd.Timestamp,
    ) -> List[Tuple[int, int]]:
        use_antig = bool(antig_sel)
        use_edad = bool(edad_sel)

        antig_list = [b for b in antig_sel if b in TENURE_BUCKETS] if use_antig else []
        edad_list = [b for b in edad_sel if b in AGE_BUCKETS] if use_edad else []
        edad_allow_sindato = use_edad and (MISSING_LABEL in edad_sel)
        dob_missing = pd.isna(dob_ts)
        if use_edad and dob_missing and not edad_allow_sindato:
            return []

        if (not use_antig) and (not use_edad):
            return [(base_s, base_e)]

        if use_antig and antig_list:
            antig_ranges = []
            for b in antig_list:
                a0, a1 = TENURE_BUCKETS[b]
                s = max(ini_day + a0, base_s)
                e = min(base_e, (base_e if a1 is None else ini_day + a1))
                if s <= e:
                    antig_ranges.append((s, e))
            if not antig_ranges:
                return []
        else:
            antig_ranges = [(base_s, base_e)]

        if use_edad and (not dob_missing) and edad_list:
            dob_date = pd.Timestamp(dob_ts).date()
            edad_ranges = []
            for b in edad_list:
                y0, y1 = AGE_BUCKETS[b]
                s_date = start.date() if y0 is None else _years_offset_date(dob_date, y0)
                e_date = end.date() if y1 is None else (_years_offset_date(dob_date, y1 + 1) - timedelta(days=1))
                s = max(_day_int(pd.Timestamp(s_date)), base_s)
                e = min(_day_int(pd.Timestamp(e_date)), base_e)
                if s <= e:
                    edad_ranges.append((s, e))
            if not edad_ranges:
                return []
        else:
            edad_ranges = [(base_s, base_e)]

        segs: List[Tuple[int, int]] = []
        for (as_, ae_) in antig_ranges:
            for (es_, ee_) in edad_ranges:
                s = max(as_, es_)
                e = min(ae_, ee_)
                if s <= e:
                    segs.append((s, e))
        return segs

    rows = []
    for _, w in windows.iterrows():
        ws = pd.Timestamp(w["window_start"]).normalize()
        we = pd.Timestamp(w["window_end"]).normalize()
        ws_day = _day_int(ws)
        we_day = _day_int(we)

        worked_cost = 0.0
        lost_cost = 0.0

        sp = spells[(spells["ini_day"] <= we_day) & (spells["fin_day"] >= ws_day)]
        if not sp.empty:
            for r in sp.itertuples(index=False):
                cod = str(r.cod)
                if cod not in cost_map:
                    continue
                ini_d = int(r.ini_day)
                fin_d = int(r.fin_day)
                base_s = max(ini_d, ws_day)
                base_e = min(fin_d, we_day)
                if base_s > base_e:
                    continue

                segs = _segments_for_spell_with_buckets(
                    ini_day=ini_d,
                    base_s=base_s,
                    base_e=base_e,
                    dob_ts=r.fnac,
                    antig_sel=antig_sel,
                    edad_sel=edad_sel,
                    start=ws,
                    end=we,
                )
                if not segs:
                    continue

                c_starts, c_ends, c_costs = cost_map[cod]
                rp = float(r.r_pct) if pd.notna(r.r_pct) else 1.0

                for (ss, ee) in segs:
                    mask = (c_starts <= ee) & (c_ends >= ss)
                    if not mask.any():
                        continue
                    idxs = np.where(mask)[0]
                    for j in idxs:
                        s2 = max(int(c_starts[j]), ss)
                        e2 = min(int(c_ends[j]), ee)
                        if s2 <= e2:
                            days = (e2 - s2 + 1)
                            worked_cost += days * float(c_costs[j]) * rp

        if not exits.empty:
            exw = exits[(exits["exit_day"] >= ws_day) & (exits["exit_day"] <= we_day)]
            if not exw.empty:
                for r in exw.itertuples(index=False):
                    cod = str(r.cod)
                    if cod not in cost_map:
                        continue
                    exit_day = int(r.exit_day)
                    ss = exit_day + 1
                    ee = we_day
                    if ss > ee:
                        continue
                    c_starts, c_ends, c_costs = cost_map[cod]
                    rp = float(r.r_pct) if pd.notna(r.r_pct) else 1.0
                    mask = (c_starts <= ee) & (c_ends >= ss)
                    if not mask.any():
                        continue
                    idxs = np.where(mask)[0]
                    for j in idxs:
                        s2 = max(int(c_starts[j]), ss)
                        e2 = min(int(c_ends[j]), ee)
                        if s2 <= e2:
                            days = (e2 - s2 + 1)
                            lost_cost += days * float(c_costs[j]) * rp

        potential = worked_cost + lost_cost
        kpi_cost = (worked_cost / potential) if potential > 0 else np.nan
        lost_rate = (lost_cost / potential) if potential > 0 else np.nan

        rows.append({
            "Periodo": w["Periodo"],
            "cut": w["cut"],
            "window_start": ws,
            "window_end": we,
            "Dias_Periodo": int((we - ws).days + 1),
            "WorkedCost": worked_cost,
            "LostCost": lost_cost,
            "PotentialCost": potential,
            "KPI_COST": kpi_cost,
            "LostRate": lost_rate,
        })

    return pd.DataFrame(rows).sort_values("cut").reset_index(drop=True)

@st.cache_data(show_spinner=False)
def compute_cost_kpis_vs_meta(
    df_cost_g: pd.DataFrame,
    df_cost_b: pd.DataFrame,
    period: str,
    yellow_min: float = 0.95,
) -> pd.DataFrame:
    if df_cost_g is None or df_cost_g.empty:
        return pd.DataFrame()

    m = df_cost_g.merge(
        df_cost_b[["Periodo", "cut", "KPI_COST"]].rename(columns={"KPI_COST": "KPI_COST_b"}),
        on=["Periodo", "cut"],
        how="left",
    )

    if period == "W":
        S = 52
    elif period == "M":
        S = 12
    elif period == "D":
        S = 365
    else:
        S = 1

    p = m["KPI_COST_b"].astype(float)
    m["Meta_COST"] = (p.shift(S) + p.shift(S + 1) + p.shift(S + 2)) / 3.0

    eps = 1e-12
    m["Indice_vs_Meta_COST"] = m["KPI_COST"] / (m["Meta_COST"] + eps)
    m["Brecha_COST"] = m["KPI_COST"] - m["Meta_COST"]

    def semaforo(ratio: float) -> str:
        if np.isnan(ratio):
            return "-"
        if ratio >= 1.0:
            return "VERDE"
        if ratio >= yellow_min:
            return "AMARILLO"
        return "ROJO"

    m["Semaforo_COST"] = m["Indice_vs_Meta_COST"].apply(semaforo)
    return m.sort_values("cut").reset_index(drop=True)


# =============================================================================
# Layout superior (control de vista + toggle filtros)
# =============================================================================
top = st.container()
with top:
    c1, c2 = st.columns([1, 2], gap="large")
    with c1:
        show_filters = st.toggle(LBL_SHOW_FILTERS, value=True)
    with c2:
        view = st.radio(LBL_VIEW_PICK, options=[VIEW_1, VIEW_2], horizontal=True, index=0)


# =============================================================================
# Layout principal (según mostrar filtros)
# =============================================================================
if show_filters:
    col_filters, col_main = st.columns([1, 3], gap="large")  # 25% / 75% aprox
else:
    col_filters = None
    col_main = st.container()


# =============================================================================
# Panel de control (solo si show_filters ON)
# =============================================================================
if show_filters:
    with col_filters:
        st.subheader(PANEL_TITLE)
        tab_p, tab_f, tab_o, tab_c = st.tabs([TAB_DATA, TAB_FILTERS, TAB_OPTIONS, TAB_COST])

        # -------------------------
        # TAB: Datos & Periodo
        # -------------------------
        with tab_p:
            uploaded = st.file_uploader(LBL_UPLOAD_MAIN, type=["xlsx", "xls", "csv"], key="uploader_hist")
            path = st.text_input(LBL_PATH_MAIN, value="", key="path_hist")

            df_raw = None
            df_cost_raw = None
            sheet_hist = None
            sheet_cost = None

            if uploaded is None and not path.strip():
                st.info(MSG_LOAD_FILE_TO_START)
                st.stop()

            try:
                if uploaded is not None:
                    if uploaded.name.lower().endswith(".csv"):
                        df_raw = read_csv_any(uploaded)
                    else:
                        xls = pd.ExcelFile(uploaded)
                        sheet_hist = st.selectbox(LBL_SHEET_MAIN, options=xls.sheet_names, index=0, key="sheet_hist_upload")
                        df_raw = read_excel_strict_hist(uploaded, sheet_hist)

                        sheet_cost = find_cost_sheet_name(xls)
                        if sheet_cost:
                            with st.expander(LBL_COST_AUTO, expanded=False):
                                st.caption(f"Hoja detectada: {sheet_cost}")
                                use_cost = st.checkbox(LBL_USE_COST_SAME, value=True, key="use_cost_same")
                                if use_cost:
                                    df_cost_raw = read_excel_any(uploaded, sheet_cost)
                        else:
                            with st.expander(LBL_COST_AUTO, expanded=False):
                                st.caption("No detecté hoja 'Costo Nominal' en este Excel.")
                                up_cost = st.file_uploader(LBL_UPLOAD_COST, type=["xlsx", "xls"], key="uploader_cost")
                                if up_cost is not None:
                                    x2 = pd.ExcelFile(up_cost)
                                    sheet_cost = st.selectbox(LBL_SHEET_COST, options=x2.sheet_names, index=0, key="sheet_cost_upload")
                                    df_cost_raw = read_excel_any(up_cost, sheet_cost)
                else:
                    p = path.strip()
                    if not os.path.exists(p):
                        st.error(MSG_PATH_NOT_FOUND)
                        st.stop()
                    if p.lower().endswith(".csv"):
                        df_raw = read_csv_any(p)
                    else:
                        xls = pd.ExcelFile(p)
                        sheet_hist = st.selectbox(LBL_SHEET_MAIN, options=xls.sheet_names, index=0, key="sheet_hist_path")
                        df_raw = read_excel_strict_hist(p, sheet_hist)

                        sheet_cost = find_cost_sheet_name(xls)
                        if sheet_cost:
                            with st.expander(LBL_COST_AUTO, expanded=False):
                                use_cost = st.checkbox(LBL_USE_COST_SAME, value=True, key="use_cost_same_path")
                                if use_cost:
                                    df_cost_raw = read_excel_any(p, sheet_cost)
                        else:
                            with st.expander(LBL_COST_AUTO, expanded=False):
                                st.caption("No detecté hoja 'Costo Nominal' en este Excel.")
                                cost_path = st.text_input(LBL_PATH_COST, value="", key="path_cost")
                                if cost_path.strip() and os.path.exists(cost_path.strip()):
                                    x2 = pd.ExcelFile(cost_path.strip())
                                    sheet_cost = st.selectbox(LBL_SHEET_COST, options=x2.sheet_names, index=0, key="sheet_cost_path")
                                    df_cost_raw = read_excel_any(cost_path.strip(), sheet_cost)

            except Exception as e:
                st.error(f"{MSG_READ_FAIL} {e}")
                st.stop()

            try:
                df0 = validate_and_prepare_hist(df_raw)
            except Exception as e:
                st.error(str(e))
                st.stop()

            df_cost = None
            if df_cost_raw is not None:
                try:
                    df_cost = validate_and_prepare_cost(df_cost_raw)
                except Exception as e:
                    st.warning(f"No pude preparar Costo Nominal: {e}")
                    df_cost = None

            df_intervals_all = merge_intervals_per_person(df0)

            min_date = df_intervals_all["ini"].min()
            max_date = df_intervals_all["fin_eff"].max()
            default_end = min(today_dt(), max_date) if pd.notna(max_date) else today_dt()
            default_start = max(min_date, default_end - pd.Timedelta(days=180)) if pd.notna(min_date) else (default_end - pd.Timedelta(days=180))

            preset = st.selectbox(
                LBL_RANGE_PRESET,
                options=["Personalizado", "Últimos 30 días", "Últimos 90 días", "Últimos 180 días", "Últimos 365 días", "Año actual (YTD)"],
                index=2,
                key="range_preset",
            )

            if "date_range_main" not in st.session_state:
                st.session_state["date_range_main"] = (default_start.date(), default_end.date())

            if preset != "Personalizado":
                end_p = default_end.date()
                if preset == "Últimos 30 días":
                    start_p = (default_end - pd.Timedelta(days=30)).date()
                elif preset == "Últimos 90 días":
                    start_p = (default_end - pd.Timedelta(days=90)).date()
                elif preset == "Últimos 180 días":
                    start_p = (default_end - pd.Timedelta(days=180)).date()
                elif preset == "Últimos 365 días":
                    start_p = (default_end - pd.Timedelta(days=365)).date()
                else:
                    start_p = date(default_end.year, 1, 1)

                if pd.notna(min_date):
                    start_p = max(start_p, min_date.date())
                if pd.notna(max_date):
                    end_p = min(end_p, max_date.date())

                st.session_state["date_range_main"] = (start_p, end_p)

            r0, r1 = st.slider(
                LBL_RANGE_SLIDER,
                min_value=(min_date.date() if pd.notna(min_date) else date(2000, 1, 1)),
                max_value=(max_date.date() if pd.notna(max_date) else default_end.date()),
                value=st.session_state["date_range_main"],
                key="date_range_slider",
            )
            st.session_state["date_range_main"] = (r0, r1)

            start_dt = pd.Timestamp(st.session_state["date_range_main"][0])
            end_dt = pd.Timestamp(st.session_state["date_range_main"][1])
            if start_dt > end_dt:
                st.error("Inicio > Fin.")
                st.stop()

            period_label = st.selectbox(LBL_GROUP_BY, options=["Día", "Semana", "Mes", "Año"], index=1, key="period_group")
            period = {"Día": "D", "Semana": "W", "Mes": "M", "Año": "Y"}[period_label]

            snap_date = st.slider(
                LBL_SNAPSHOT_DATE,
                min_value=start_dt.date(),
                max_value=end_dt.date(),
                value=end_dt.date(),
                key="snap_date",
            )
            snap_dt = pd.Timestamp(snap_date)

            cut_today = min(today_dt(), max_date) if pd.notna(max_date) else today_dt()
            st.write(f"{LBL_TODAY_CUT}: **{cut_today.date()}**")

            st.session_state["__globals__"] = {
                "df0": df0,
                "df_cost": df_cost,
                "df_intervals_all": df_intervals_all,
                "start_dt": start_dt,
                "end_dt": end_dt,
                "period": period,
                "period_label": period_label,
                "snap_dt": snap_dt,
                "cut_today": cut_today,
            }

        # -------------------------
        # TAB: Filtros
        # -------------------------
        with tab_f:
            g = st.session_state.get("__globals__")
            if not g:
                st.stop()
            df0 = g["df0"]

            st.caption(LBL_FILTERS_HINT)

            def opts(df: pd.DataFrame, col: str) -> List[str]:
                v = df[col].dropna().astype(str).str.strip()
                v = v[v != ""].unique().tolist()
                return sorted(v)

            if st.button(BTN_CLEAR_FILTERS, use_container_width=True, key="btn_clear_filters"):
                for k in [
                    "f_sexo", "f_area_gen", "f_area", "f_cargo", "f_clas", "f_ts", "f_emp", "f_nac", "f_lug", "f_reg",
                    "f_antig", "f_edad",
                ]:
                    st.session_state[k] = []
                st.rerun()

            area_gen_pick = st.multiselect(LBL_AREA_GEN, opts(df0, "area_gen"), default=st.session_state.get("f_area_gen", []), key="f_area_gen")

            if area_gen_pick:
                df_area = df0[df0["area_gen"].isin(area_gen_pick)]
                area_opts = opts(df_area, "area")
            else:
                area_opts = opts(df0, "area")

            fs = FilterState(
                sexo=st.multiselect(LBL_SEXO, opts(df0, "sexo"), default=st.session_state.get("f_sexo", []), key="f_sexo"),
                area_gen=area_gen_pick,
                area=st.multiselect(LBL_AREA, area_opts, default=st.session_state.get("f_area", []), key="f_area"),
                cargo=st.multiselect(LBL_CARGO, opts(df0, "cargo"), default=st.session_state.get("f_cargo", []), key="f_cargo"),
                clas=st.multiselect(LBL_CLAS, opts(df0, "clas"), default=st.session_state.get("f_clas", []), key="f_clas"),
                ts=st.multiselect(LBL_TS, opts(df0, "ts"), default=st.session_state.get("f_ts", []), key="f_ts"),
                emp=st.multiselect(LBL_EMP, opts(df0, "emp"), default=st.session_state.get("f_emp", []), key="f_emp"),
                nac=st.multiselect(LBL_NAC, opts(df0, "nac"), default=st.session_state.get("f_nac", []), key="f_nac"),
                lug=st.multiselect(LBL_LUG, opts(df0, "lug"), default=st.session_state.get("f_lug", []), key="f_lug"),
                reg=st.multiselect(LBL_REG, opts(df0, "reg"), default=st.session_state.get("f_reg", []), key="f_reg"),
                antig=st.multiselect(
                    LBL_TENURE_BUCKET,
                    list(TENURE_BUCKETS.keys()) + [MISSING_LABEL],
                    default=st.session_state.get("f_antig", []),
                    key="f_antig",
                ),
                edad=st.multiselect(
                    LBL_AGE_BUCKET,
                    list(AGE_BUCKETS.keys()) + [MISSING_LABEL],
                    default=st.session_state.get("f_edad", []),
                    key="f_edad",
                ),
            )
            st.session_state["__fs__"] = fs

        # -------------------------
        # TAB: Opciones
        # -------------------------
        with tab_o:
            unique_personas_por_dia = st.checkbox(LBL_OPT_UNIQUE_DAY, value=True, key="opt_unique_day")
            show_labels = st.checkbox(LBL_OPT_SHOW_LABELS, value=False, key="opt_show_labels")

            st.markdown(f"**{LBL_OPT_HORIZON}**")
            h_choice = st.selectbox(LBL_OPT_H_CHOICE, options=["30 días", "90 días", "180 días", "Otro…"], index=0, key="opt_h_choice")
            if h_choice == "30 días":
                horizon_days = 30
            elif h_choice == "90 días":
                horizon_days = 90
            elif h_choice == "180 días":
                horizon_days = 180
            else:
                horizon_days = int(st.number_input(LBL_OPT_H_CUSTOM, min_value=7, max_value=365, value=30, step=1, key="opt_h_custom"))

            st.markdown(f"**{LBL_OPT_SEMAFORO}**")
            green_max = float(st.number_input(LBL_OPT_GREEN, min_value=0.70, max_value=1.10, value=0.95, step=0.01, key="opt_green"))
            yellow_max = float(st.number_input(LBL_OPT_YELLOW, min_value=0.80, max_value=1.30, value=1.05, step=0.01, key="opt_yellow"))

            st.markdown(f"**{LBL_OPT_CONT}**")
            top_rows = int(st.number_input(LBL_OPT_TOP_ROWS, min_value=10, max_value=200, value=40, step=5, key="opt_top_rows"))
            top_cols = int(st.number_input(LBL_OPT_TOP_COLS, min_value=10, max_value=200, value=40, step=5, key="opt_top_cols"))

            st.session_state["__opts__"] = {
                "unique_personas_por_dia": unique_personas_por_dia,
                "show_labels": show_labels,
                "horizon_days": horizon_days,
                "green_max": green_max,
                "yellow_max": yellow_max,
                "top_rows": top_rows,
                "top_cols": top_cols,
            }

        # -------------------------
        # TAB: Costo KPI
        # -------------------------
        with tab_c:
            st.caption(LBL_COST_HINT)
            yellow_min_cost = float(st.number_input(LBL_COST_YELLOW_MIN, min_value=0.50, max_value=0.99, value=0.95, step=0.01, key="opt_cost_yellow"))
            st.session_state["__cost_opts__"] = {"yellow_min_cost": yellow_min_cost}


# =============================================================================
# Main (si filtros ocultos, requiere que ya exista __globals__)
# =============================================================================
with (col_main if hasattr(col_main, "__enter__") else st.container()):
    g = st.session_state.get("__globals__")
    fs = st.session_state.get("__fs__")
    opts = st.session_state.get("__opts__")
    cost_opts = st.session_state.get("__cost_opts__", {})

    if not g or not fs or not opts:
        st.warning(MSG_NEED_DATA if not show_filters else MSG_LOAD_FILE_TO_START)
        st.stop()

    df0 = g["df0"]
    df_cost = g["df_cost"]
    df_intervals_all = g["df_intervals_all"]
    start_dt = g["start_dt"]
    end_dt = g["end_dt"]
    period = g["period"]
    period_label = g["period_label"]
    snap_dt = g["snap_dt"]
    cut_today = g["cut_today"]

    unique_personas_por_dia = bool(opts["unique_personas_por_dia"])
    show_labels = bool(opts["show_labels"])
    horizon_days = int(opts["horizon_days"])
    green_max = float(opts["green_max"])
    yellow_max = float(opts["yellow_max"])
    yellow_min_cost = float(cost_opts.get("yellow_min_cost", 0.95))

    # Aplicar filtros categóricos
    df0_f = apply_categorical_filters(df0, fs)

    # Tenure bins para RR
    rr_bins = [(k, TENURE_BUCKETS[k][0], TENURE_BUCKETS[k][1]) for k in TENURE_BUCKETS.keys()]

    # -------------------------
    # CALC: RR (grupo + baseline)
    # -------------------------
    with st.spinner("Calculando Rotación robusta (persona-día) + meta estacional..."):
        rr_g = compute_rr_period_table(
            df_events=df0_f,
            start=start_dt,
            end=end_dt,
            period=period,
            tenure_bins=rr_bins,
            antig_sel=[x for x in fs.antig if x in TENURE_BUCKETS] if fs.antig else [],
            edad_sel=fs.edad,
            unique_salidas_por_dia=unique_personas_por_dia,
        )
        rr_b = compute_rr_period_table(
            df_events=df0,
            start=start_dt,
            end=end_dt,
            period=period,
            tenure_bins=rr_bins,
            antig_sel=[x for x in fs.antig if x in TENURE_BUCKETS] if fs.antig else [],
            edad_sel=fs.edad,
            unique_salidas_por_dia=unique_personas_por_dia,
        )
        rr_kpi = compute_rr_kpi(rr_g, rr_b, period=period, green_max=green_max, yellow_max=yellow_max)

    # -------------------------
    # CALC: KM supervivencia (TOTAL)
    # -------------------------
    with st.spinner("Calculando supervivencia (KM simple)..."):
        # KM sobre spells del grupo (ya con filtros categóricos); edad/antig buckets no se aplican aquí para no complicar censoring
        km_curve, surv_H = km_survival_curve(df0_f, start_dt, end_dt, horizon_days)

    # -------------------------
    # Vista 1: Rotación y Supervivencia
    # -------------------------
    if view == VIEW_1:
        st.subheader(VIEW_1)

        if rr_kpi.empty:
            st.warning(MSG_NO_DATA_FOR_VIEW)
            st.stop()

        rr_total = rr_kpi[rr_kpi["Bin"] == "TOTAL"].copy().sort_values("cut")
        rr_total_nonnull = rr_total.dropna(subset=["RateW_1000"])

        # Último periodo visible
        if rr_total_nonnull.empty:
            last_row = rr_total.iloc[-1]
        else:
            last_row = rr_total_nonnull.iloc[-1]

        last_ratew = float(last_row["RateW_1000"]) if pd.notna(last_row["RateW_1000"]) else np.nan
        last_sema = str(last_row.get("Semaforo", "-"))
        last_exits_w = float(last_row["Exits_w"]) if pd.notna(last_row["Exits_w"]) else 0.0

        # Row 1: KPIs Ejecutivos
        st.markdown(f"### {V1_ROW1_KPI_TITLE}")
        k1, k2, k3 = st.columns([1, 1, 1], gap="large")
        with k1:
            st.metric(KPI1_TITLE, "-" if np.isnan(last_ratew) else f"{last_ratew:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."),
                      help=f"Semáforo: {last_sema}")
            st.caption(f"Semáforo: **{last_sema}**")
        with k2:
            st.metric(KPI2_TITLE, "-" if np.isnan(surv_H) else f"{surv_H*100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", "."),
                      help="Cohorte: spells con ini dentro del rango seleccionado.")
        with k3:
            st.metric(KPI3_TITLE, f"{last_exits_w:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."),
                      help="Suma ponderada (r_pct) de salidas en el último periodo visible.")

        # Row 2: tendencia + supervivencia
        cL, cR = st.columns([2, 1], gap="large")
        with cL:
            st.markdown(f"### {V1_ROW2_LEFT_TITLE}")
            fig = px.line(rr_total, x="Periodo", y=["RateW_1000", "Meta"], title=V1_ROW2_LEFT_TITLE)
            fig = nice_xaxis(fig)
            st.plotly_chart(fig, use_container_width=True)

        with cR:
            st.markdown(f"### {V1_ROW2_RIGHT_TITLE}")
            if km_curve.empty:
                st.info(MSG_NO_DATA_FOR_VIEW)
            else:
                figk = px.line(km_curve, x="day", y="S", title=V1_ROW2_RIGHT_TITLE)
                figk.update_yaxes(range=[0, 1])
                st.plotly_chart(figk, use_container_width=True)

        # Row 3: desglose por Área y Cargo (último periodo)
        st.markdown(f"### {V1_ROW3_TITLE}")

        # Elegir periodo a desglosar (default último)
        period_choices = rr_total["Periodo"].astype(str).tolist()
        default_idx = len(period_choices) - 1 if period_choices else 0
        pick_period = st.selectbox("Periodo a desglosar", options=period_choices, index=default_idx, key="pick_period_breakdown")

        # Window bounds del periodo elegido
        wrow = rr_total[rr_total["Periodo"].astype(str) == str(pick_period)].iloc[0]
        ws = pd.Timestamp(wrow["window_start"]).normalize()
        we = pd.Timestamp(wrow["window_end"]).normalize()

        # Exposure + exits_w por (area, cargo) en ese periodo
        # (reusa lógica persona-día sin expandir)
        @st.cache_data(show_spinner=False)
        def rr_breakdown_area_cargo(
            df_events: pd.DataFrame,
            ws: pd.Timestamp,
            we: pd.Timestamp,
            antig_sel: List[str],
            edad_sel: List[str],
            tenure_bins: List[Tuple[str, int, Optional[int]]],
        ) -> pd.DataFrame:
            if df_events.empty:
                return pd.DataFrame()

            ws = pd.Timestamp(ws).normalize()
            we = pd.Timestamp(we).normalize()
            ws_day = _day_int(ws)
            we_day = _day_int(we)

            # bins elegidos (si antig_sel)
            allowed_bins = [b for b, _, _ in tenure_bins]
            if antig_sel:
                allowed_bins = [b for b in allowed_bins if b in set(antig_sel)]
            tbins = [t for t in tenure_bins if t[0] in allowed_bins]

            dfe = df_events[(df_events["ini"] <= we) & (df_events["fin_eff"] >= ws)].copy()
            if dfe.empty:
                return pd.DataFrame()

            dfe["ini_day"] = dfe["ini"].apply(_day_int).astype(np.int64)
            dfe["fin_day"] = dfe["fin_eff"].apply(_day_int).astype(np.int64)

            expo: Dict[Tuple[str, str], float] = {}

            for r in dfe.itertuples(index=False):
                ini_day = int(r.ini_day)
                fin_day = int(r.fin_day)
                base_s = max(ini_day, ws_day)
                base_e = min(fin_day, we_day)
                if base_s > base_e:
                    continue

                age_allowed = _age_allowed_segments(r.fnac, ws_day, we_day, edad_sel)
                if edad_sel and not age_allowed:
                    continue

                base_seg = (base_s, base_e)

                # acumula exposure sumando bins permitidos
                segs_by_bin = _tenure_bin_segments_for_spell(ini_day, base_seg, tbins)
                exp_len = 0
                for seg in segs_by_bin.values():
                    exp_len += _intersect_len(seg, age_allowed) if edad_sel else (seg[1] - seg[0] + 1)

                if exp_len <= 0:
                    continue

                key = (str(r.area), str(r.cargo))
                expo[key] = expo.get(key, 0.0) + float(exp_len)

            exp_df = pd.DataFrame(
                [{"area": k[0], "cargo": k[1], "Exposure": v} for k, v in expo.items()]
            )

            # exits_w por area,cargo en el periodo
            exits = df_events[~df_events["fin"].isna()].copy()
            exits = exits[(exits["fin"] >= ws) & (exits["fin"] <= we)].copy()
            if not exits.empty:
                exits["tenure_days_exit"] = (exits["fin"] - exits["ini"]).dt.days.astype("Int64")
                exits = exits[~exits["tenure_days_exit"].isna()].copy()
                exits["tenure_days_exit"] = exits["tenure_days_exit"].astype(int)

                exits["Edad"] = bucket_edad_from_dob(exits["fnac"], exits["fin"])
                if edad_sel:
                    if MISSING_LABEL in edad_sel:
                        exits = exits[exits["Edad"].isin(edad_sel)]
                    else:
                        exits = exits[(exits["Edad"].isin(edad_sel)) & (exits["Edad"] != MISSING_LABEL)]

                exits["Bin"] = exits["tenure_days_exit"].apply(lambda x: _assign_tenure_bin(int(x), tbins))
                exits = exits[~exits["Bin"].isna()].copy()
                exits = exits[exits["Bin"].isin([t[0] for t in tbins])].copy()

                ex_agg = exits.groupby(["area", "cargo"], as_index=False).agg(
                    Exits=("cod", "size"),
                    Exits_w=("r_pct", "sum"),
                )
            else:
                ex_agg = pd.DataFrame(columns=["area", "cargo", "Exits", "Exits_w"])

            out = exp_df.merge(ex_agg, on=["area", "cargo"], how="left")
            out["Exits"] = out["Exits"].fillna(0).astype(float)
            out["Exits_w"] = out["Exits_w"].fillna(0.0).astype(float)
            out["RateW_1000"] = np.where(out["Exposure"] > 0, (out["Exits_w"] / out["Exposure"]) * 1000.0, np.nan)
            return out.sort_values("RateW_1000", ascending=False).reset_index(drop=True)

        bdf = rr_breakdown_area_cargo(
            df_events=df0_f,
            ws=ws,
            we=we,
            antig_sel=[x for x in fs.antig if x in TENURE_BUCKETS] if fs.antig else [],
            edad_sel=fs.edad,
            tenure_bins=rr_bins,
        )

        if bdf.empty:
            st.info(MSG_NO_DATA_FOR_VIEW)
        else:
            st.dataframe(_safe_table_for_streamlit(bdf), use_container_width=True, height=420)

        # -------------------------
        # Legacy: PDE + KPI_COST (no romper)
        # -------------------------
        with st.expander(V1_LEGACY_EXPANDER, expanded=False):
            with st.spinner("Calculando PDE/KPI_COST (legacy)..."):
                df_intervals_f = merge_intervals_per_person(df0_f) if not df0_f.empty else df0_f.copy()

                df_salidas_daily_g, df_sal_det = compute_salidas_daily_filtered(
                    df_events=df0_f,
                    start=start_dt,
                    end=end_dt,
                    antig_sel=fs.antig,
                    edad_sel=fs.edad,
                    unique_personas_por_dia=unique_personas_por_dia,
                )
                df_exist_daily_g = compute_existencias_daily_filtered_fast(
                    df_intervals=df_intervals_f,
                    start=start_dt,
                    end=end_dt,
                    antig_sel=fs.antig,
                    edad_sel=fs.edad,
                )
                df_daily_g = df_salidas_daily_g.merge(df_exist_daily_g[["Día", "Existencias"]], on="Día", how="left")
                df_daily_g["Existencias"] = df_daily_g["Existencias"].fillna(0).astype(int)
                df_daily_g = add_calendar_fields(df_daily_g, "Día")

                df_salidas_daily_b, _ = compute_salidas_daily_filtered(
                    df_events=df0,
                    start=start_dt,
                    end=end_dt,
                    antig_sel=fs.antig,
                    edad_sel=fs.edad,
                    unique_personas_por_dia=unique_personas_por_dia,
                )
                df_exist_daily_b = compute_existencias_daily_filtered_fast(
                    df_intervals=df_intervals_all,
                    start=start_dt,
                    end=end_dt,
                    antig_sel=fs.antig,
                    edad_sel=fs.edad,
                )
                df_daily_b = df_salidas_daily_b.merge(df_exist_daily_b[["Día", "Existencias"]], on="Día", how="left")
                df_daily_b["Existencias"] = df_daily_b["Existencias"].fillna(0).astype(int)
                df_daily_b = add_calendar_fields(df_daily_b, "Día")

                df_period_g = aggregate_daily_to_period_for_pde(df_daily_g, period)
                df_period_b = aggregate_daily_to_period_for_pde(df_daily_b, period)

                df_kpi_pde = compute_pde_kpis(
                    df_period_g=df_period_g,
                    df_period_base=df_period_b,
                    horizon_days=horizon_days,
                    period=period,
                    green_max=green_max,
                    yellow_max=yellow_max,
                )

                st.markdown(f"### {V1_LEGACY_PDE_TITLE}")
                if df_kpi_pde.empty or df_kpi_pde["KPI_PDE"].dropna().empty:
                    st.info(MSG_NO_DATA_FOR_VIEW)
                else:
                    figp = px.line(df_kpi_pde, x="Periodo", y=["KPI_PDE"], title=V1_LEGACY_PDE_TITLE)
                    st.plotly_chart(nice_xaxis(figp), use_container_width=True)

                st.markdown(f"### {V1_LEGACY_COST_TITLE}")
                df_kpi_cost = pd.DataFrame()
                if df_cost is not None and not df_cost.empty:
                    df_cost_g = compute_cost_period_metrics(
                        df_events=df0_f,
                        df_cost=df_cost,
                        start=start_dt,
                        end=end_dt,
                        period=period,
                        antig_sel=fs.antig,
                        edad_sel=fs.edad,
                        unique_salidas_por_dia=unique_personas_por_dia,
                    )
                    df_cost_b = compute_cost_period_metrics(
                        df_events=df0,
                        df_cost=df_cost,
                        start=start_dt,
                        end=end_dt,
                        period=period,
                        antig_sel=fs.antig,
                        edad_sel=fs.edad,
                        unique_salidas_por_dia=unique_personas_por_dia,
                    )
                    df_kpi_cost = compute_cost_kpis_vs_meta(df_cost_g, df_cost_b, period=period, yellow_min=yellow_min_cost)

                if df_kpi_cost.empty:
                    st.info("Costo KPI desactivado (no hay hoja Costo Nominal o no pudo leerse).")
                else:
                    figc = px.line(df_kpi_cost, x="Periodo", y=["KPI_COST", "Meta_COST"], title=V1_LEGACY_COST_TITLE)
                    st.plotly_chart(nice_xaxis(figc), use_container_width=True)

                # Descarga legacy historia
                buf_xlsx = io.BytesIO()
                with pd.ExcelWriter(buf_xlsx, engine="openpyxl") as writer:
                    df_daily_g.to_excel(writer, index=False, sheet_name="Diario_Grupo")
                    df_period_g.to_excel(writer, index=False, sheet_name="Periodo_PDE_Grupo")
                    df_kpi_pde.to_excel(writer, index=False, sheet_name="KPI_PDE")
                    df_sal_det.to_excel(writer, index=False, sheet_name="Salidas_Detalle")
                    rr_kpi.to_excel(writer, index=False, sheet_name="RR_Rotation")
                    if not df_kpi_cost.empty:
                        df_kpi_cost.to_excel(writer, index=False, sheet_name="KPI_COST")
                st.download_button(
                    DL_HISTORY_XLSX,
                    data=buf_xlsx.getvalue(),
                    file_name=FILE_HISTORY_XLSX,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="dl_historia",
                )

    # -------------------------
    # Vista 2: Existencias Actuales
    # -------------------------
    else:
        st.subheader(VIEW_2)

        # Snapshot activos hoy (con filtros categóricos y buckets)
        df_now = df0_f[(df0_f["ini"] <= cut_today) & (df0_f["fin_eff"] >= cut_today)].copy()
        if not df_now.empty:
            df_now["ref"] = cut_today
            df_now["antig_dias"] = (df_now["ref"] - df_now["ini"]).dt.days
            df_now["Antigüedad"] = bucket_antiguedad(df_now["antig_dias"])
            df_now["Edad"] = bucket_edad_from_dob(df_now["fnac"], df_now["ref"])
            if fs.antig:
                df_now = df_now[df_now["Antigüedad"].isin(fs.antig)]
            if fs.edad:
                df_now = df_now[df_now["Edad"].isin(fs.edad)]

        # tabla por persona activa (antigüedad desde último ingreso)
        df_all = df0_f.copy()
        df_all["fin_n"] = df_all["fin_eff"].clip(upper=cut_today)
        df_all["days_spell"] = (df_all["fin_n"] - df_all["ini"]).dt.days + 1
        df_all["days_spell"] = df_all["days_spell"].clip(lower=0).fillna(0).astype(int)

        per = df_all.groupby("cod", as_index=False).agg(
            Ultimo_Ingreso=("ini", "max"),
            Antig_Acumulada_Dias=("days_spell", "sum"),
            r_pct=("r_pct", "last"),
            area_gen=("area_gen", "last"),
            area=("area", "last"),
            cargo=("cargo", "last"),
            clas=("clas", "last"),
            sexo=("sexo", "last"),
        )
        per["Antig_UltimoIngreso_Dias"] = (cut_today - per["Ultimo_Ingreso"]).dt.days

        activos_hoy = set(df_now["cod"].unique().tolist()) if not df_now.empty else set()
        per_act = per[per["cod"].isin(activos_hoy)].copy()

        # Row 1 KPIs
        exist_hoy = int(len(activos_hoy))
        avg_r = float(per_act["r_pct"].mean()) if not per_act.empty else np.nan
        avg_ten = float(per_act["Antig_UltimoIngreso_Dias"].mean()) if not per_act.empty else np.nan

        k1, k2, k3 = st.columns(3, gap="large")
        k1.metric(V2_ROW1_KPI1, f"{exist_hoy:,}".replace(",", "."))
        k2.metric(V2_ROW1_KPI2, "-" if np.isnan(avg_r) else f"{avg_r:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
        k3.metric(V2_ROW1_KPI3, "-" if np.isnan(avg_ten) else f"{avg_ten:,.1f}".replace(",", "X").replace(".", ",").replace("X", "."))

        # Row 2 charts
        cL, cR = st.columns([1, 1], gap="large")
        with cL:
            st.markdown(f"### {V2_ROW2_LEFT_TITLE}")
            if per_act.empty:
                st.info(MSG_NO_DATA_FOR_VIEW)
            else:
                clas_counts = per_act["clas"].fillna(MISSING_LABEL).astype(str).value_counts().reset_index()
                clas_counts.columns = ["Clasificación", "Activos"]
                figc = px.bar(
                    clas_counts.sort_values("Activos"),
                    x="Activos",
                    y="Clasificación",
                    orientation="h",
                    title=V2_ROW2_LEFT,
                )
                figc = apply_bar_labels(figc, show_labels, fmt=".0f")
                st.plotly_chart(figc, use_container_width=True)

        with cR:
            st.markdown(f"### {V2_ROW2_RIGHT}")
            if per_act.empty:
                st.info(MSG_NO_DATA_FOR_VIEW)
            else:
                # Bucket edad en snapshot (usa fnac vs cut_today)
                per_act2 = per_act.copy()
                # Necesitamos DOB por persona: tomamos del df0_f (último registro)
                dob_map = df0_f.sort_values(["cod", "ini"]).groupby("cod")["fnac"].last()
                per_act2["fnac"] = per_act2["cod"].map(dob_map)
                per_act2["EdadBucket"] = bucket_edad_from_dob(per_act2["fnac"], pd.Series([cut_today] * len(per_act2), index=per_act2.index))
                # si el usuario ya filtró edad buckets, mantenemos consistencia
                if fs.edad:
                    per_act2 = per_act2[per_act2["EdadBucket"].isin(fs.edad)]
                age_counts = per_act2["EdadBucket"].fillna(MISSING_LABEL).astype(str).value_counts().reset_index()
                age_counts.columns = ["Edad (bucket)", "Activos"]
                figa = px.bar(
                    age_counts.sort_values("Activos"),
                    x="Activos",
                    y="Edad (bucket)",
                    orientation="h",
                    title=V2_ROW2_RIGHT,
                )
                figa = apply_bar_labels(figa, show_labels, fmt=".0f")
                st.plotly_chart(figa, use_container_width=True)

        # Row 3 tabla completa
        st.markdown(f"### {V2_ROW3_TITLE}")
        if per_act.empty:
            st.info(MSG_NO_DATA_FOR_VIEW)
        else:
            show_tbl = per_act.copy()
            show_tbl = show_tbl.rename(columns={
                "area_gen": "Área General",
                "area": "Área",
                "cargo": "Cargo",
                "clas": "Clasificación",
                "sexo": "Sexo",
                "r_pct": "r_pct",
                "Antig_UltimoIngreso_Dias": "Antigüedad (días)",
                "Ultimo_Ingreso": "Último ingreso",
            })
            show_tbl = show_tbl[[
                "Área General", "Área", "Cargo", "Clasificación", "Sexo",
                "Antigüedad (días)", "Último ingreso", "r_pct", "cod"
            ]].sort_values(["Área General", "Área", "Cargo", "Antigüedad (días)"], ascending=[True, True, True, False])

            st.dataframe(_safe_table_for_streamlit(show_tbl), use_container_width=True, height=520)

            # descarga snapshot
            buf2 = io.BytesIO()
            with pd.ExcelWriter(buf2, engine="openpyxl") as writer:
                show_tbl.to_excel(writer, index=False, sheet_name="Activos_Snapshot")
                per_act.to_excel(writer, index=False, sheet_name="Activos_Detalle")
            st.download_button(
                DL_CURRENT_XLSX,
                data=buf2.getvalue(),
                file_name=FILE_CURRENT_XLSX,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_actual",
            )

