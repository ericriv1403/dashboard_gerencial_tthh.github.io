# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import os
from dataclasses import dataclass
from datetime import date, timedelta
from typing import List, Tuple, Optional, Dict

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# =============================================================================
# Config
# =============================================================================
st.set_page_config(
    page_title="RRHH | Existencias, Salidas, KPI (PDE y Costo)",
    layout="wide",
)

st.title("Panel RRHH: Existencias, Salidas y KPIs (PDE + Costo)")


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

# posibles nombres para %R (se detecta automáticamente)
R_COL_CANDIDATES = [
    "%R", "% R", "R", "R%", "Porcentaje R", "PorcentajeR", "Factor R", "FactorR"
]

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
COST_REQUIRED = [
    "Codigo Personal",
    "Fecha Inicio Corte",
    "Fecha Fin Corte",
    "Costo Nominal Diario",
]

COST_COL_MAP = {
    "Codigo Personal": "cod",
    "Fecha Inicio Corte": "c_ini",
    "Fecha Fin Corte": "c_fin",
    "Costo Nominal Diario": "costo",
}


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
# Helpers base
# =============================================================================
def _to_datetime(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce").dt.normalize()

def today_dt() -> pd.Timestamp:
    return pd.Timestamp(date.today())

def excel_weeknum_return_type_1(d: pd.Series) -> pd.Series:
    return d.dt.strftime("%U").astype(int) + 1

def week_end_sun_to_sat(d: pd.Series) -> pd.Series:
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
# Mapping Área y Clasificación
# =============================================================================
def _map_area(area_raw: pd.Series) -> Tuple[pd.Series, pd.Series]:
    key = area_raw.astype("string").str.strip()
    key_u = key.str.upper()

    std = key_u.map(lambda x: AREA_REF.get(x, (None, None))[0] if pd.notna(x) else None)
    gen = key_u.map(lambda x: AREA_REF.get(x, (None, None))[1] if pd.notna(x) else None)

    std = std.fillna(key)
    gen = gen.fillna(pd.NA)

    std = std.replace({"": pd.NA}).fillna(MISSING_LABEL).astype("string")
    gen = gen.replace({"": pd.NA}).fillna(MISSING_LABEL).astype("string")
    return std, gen

def _map_clas(clas_raw: pd.Series) -> pd.Series:
    key = clas_raw.astype("string").str.strip()
    key_u = key.str.upper()

    std = key_u.map(lambda x: CLAS_REF.get(x, None) if pd.notna(x) else None)
    std = std.fillna(key)

    std = std.replace({"": pd.NA}).fillna(MISSING_LABEL).astype("string")
    return std


# =============================================================================
# Detección/lectura robusta
# =============================================================================
def read_excel_any(file_obj_or_path, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(file_obj_or_path, sheet_name=sheet_name)

def read_excel_strict_hist(file_obj_or_path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(file_obj_or_path, sheet_name=sheet_name)
    # intentamos recortar a required + %R si existe
    cols = list(df.columns)
    r_col = next((c for c in cols if str(c).strip() in R_COL_CANDIDATES), None)
    keep = [c for c in REQUIRED_COLS if c in cols]
    if r_col and r_col not in keep:
        keep.append(r_col)
    return df[keep].copy() if keep else df.copy()

def read_csv_any(file_obj_or_path) -> pd.DataFrame:
    return pd.read_csv(file_obj_or_path)

def find_cost_sheet_name(xls: pd.ExcelFile) -> Optional[str]:
    # busca una hoja que se llame "Costo Nominal" (case-insensitive, flexible)
    for s in xls.sheet_names:
        if str(s).strip().lower() == "costo nominal":
            return s
    # fallback: contiene "costo" y "nominal"
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
    # valida columnas base (REQUIRED_COLS deben existir; %R opcional)
    missing = [c for c in REQUIRED_COLS if c not in df_raw.columns]
    if missing:
        raise ValueError("Faltan columnas requeridas en Historia Personal:\n- " + "\n- ".join(missing))

    # detecta %R si existe
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

    # %R
    if r_col:
        out = out.rename(columns={r_col: "r_pct"})
    else:
        out["r_pct"] = 1.0

    # limpia strings
    for c in ["cod", "clas_raw", "sexo", "ts", "emp", "area_raw", "cargo", "nac", "lug", "reg"]:
        out[c] = out[c].astype("string").str.strip()
        out.loc[out[c].isin(["", "None", "nan", "NaT"]), c] = pd.NA

    out = out[~out["cod"].isna()].copy()
    out = out[~out["ini"].isna()].copy()
    out["cod"] = out["cod"].astype(str)

    # r_pct numérico: si viene como 15% o 0.15, normalizamos a factor (0..1)
    rp = out["r_pct"].copy()
    if rp.dtype == "object" or str(rp.dtype).startswith("string"):
        rp2 = rp.astype(str).str.replace("%", "", regex=False).str.strip()
        rp_num = pd.to_numeric(rp2, errors="coerce")
        # si parece estar 0..100, lo bajamos a 0..1
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

    # sanity: c_fin >= c_ini
    d = d[d["c_fin"] >= d["c_ini"]].copy()
    return d


# =============================================================================
# Intervalos por persona (para existencias globales)
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
    antig: List[str]  # bucket
    edad: List[str]   # bucket

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
# Buckets
# =============================================================================
ANTIG_BUCKETS = {
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


# =============================================================================
# Existencias diarias (rápido) con buckets
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

    antig_list = [b for b in antig_sel if b in ANTIG_BUCKETS] if use_antig else []
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
                a0, a1 = ANTIG_BUCKETS[b]
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


# =============================================================================
# Salidas diarias (por fin) + buckets en fecha de salida
# =============================================================================
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


# =============================================================================
# Ingresos acumulados (para vista Actual)
# =============================================================================
def compute_ingresos_total(
    df_events: pd.DataFrame,
    start: pd.Timestamp,
    end: pd.Timestamp,
    unique_personas_en_rango: bool = True,
) -> int:
    d = df_events[~df_events["ini"].isna()].copy()
    d = d[(d["ini"] >= start) & (d["ini"] <= end)]
    if d.empty:
        return 0
    return int(d["cod"].nunique()) if unique_personas_en_rango else int(len(d))


# =============================================================================
# Period windows
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


# =============================================================================
# Agregación diaria -> periodo para PDE (ya la tienes, la mantenemos)
# =============================================================================
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

        sem = int(g["Semana"].max()) if "Semana" in g.columns else np.nan
        mes = int(g["Mes"].max()) if "Mes" in g.columns else np.nan
        anio = int(g["Año"].max()) if "Año" in g.columns else np.nan
        codsem = str(g["CodSem"].max()) if "CodSem" in g.columns else None
        codmes = str(g["CodMes"].max()) if "CodMes" in g.columns else None
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
            "Semana": sem if pd.notna(sem) else None,
            "Mes": mes if pd.notna(mes) else None,
            "Año": anio if pd.notna(anio) else None,
            "CodSem": codsem,
            "CodMes": codmes,
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


# =============================================================================
# KPI PDE (se mantiene)
# =============================================================================
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
    m["S"] = S

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

    rojo = m["KPI_PDE"] > yellow_max
    amarillo = (m["KPI_PDE"] > green_max) & (m["KPI_PDE"] <= yellow_max)

    m["Alerta_Rojo_2Seg"] = rojo & rojo.shift(1).fillna(False)
    m["Alerta_Amarillo_3Alza"] = (
        amarillo & amarillo.shift(1).fillna(False) & amarillo.shift(2).fillna(False) &
        (m["KPI_PDE"] > m["KPI_PDE"].shift(1)) & (m["KPI_PDE"].shift(1) > m["KPI_PDE"].shift(2))
    )

    m["Accion"] = np.where(
        m["Alerta_Rojo_2Seg"] | m["Alerta_Amarillo_3Alza"],
        "INTERVENIR",
        ""
    )

    def _lectura(k: float) -> str:
        if np.isnan(k):
            return "-"
        if k < 1.0:
            return "MEJOR que la meta (menos capacidad perdida)"
        if k > 1.0:
            return "PEOR que la meta (más capacidad perdida)"
        return "IGUAL a la meta"

    m["Lectura"] = m["KPI_PDE"].apply(_lectura)
    return m.sort_values("cut").reset_index(drop=True)


# =============================================================================
# KPI COSTO (nuevo): KPI_COST = 1 - (LostCost / (WorkedCost + LostCost))
# - WorkedCost: costo diario * %R para días trabajados (existencia) dentro de la ventana
# - LostCost: por cada salida en día d dentro de la ventana, costo diario * %R para días (d+1..fin_ventana)
# - Meta estacional (baseline) igual que PDE: promedio t-S, t-S-1, t-S-2
# =============================================================================
def _day_int(ts: pd.Timestamp) -> int:
    return int(np.datetime64(pd.Timestamp(ts).normalize(), "D").astype("int64"))

def _date_from_day_int(x: int) -> pd.Timestamp:
    return pd.Timestamp(np.datetime64(x, "D"))

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
    """Retorna lista de segmentos [s,e] (int days, inclusivo) elegibles según buckets (antig/edad)."""
    use_antig = bool(antig_sel)
    use_edad = bool(edad_sel)

    antig_list = [b for b in antig_sel if b in ANTIG_BUCKETS] if use_antig else []
    edad_list = [b for b in edad_sel if b in AGE_BUCKETS] if use_edad else []
    edad_allow_sindato = use_edad and (MISSING_LABEL in edad_sel)
    dob_missing = pd.isna(dob_ts)
    if use_edad and dob_missing and not edad_allow_sindato:
        return []

    # base
    if (not use_antig) and (not use_edad):
        return [(base_s, base_e)]

    # Antig ranges
    if use_antig and antig_list:
        antig_ranges = []
        for b in antig_list:
            a0, a1 = ANTIG_BUCKETS[b]
            s = max(ini_day + a0, base_s)
            e = min(base_e, (base_e if a1 is None else ini_day + a1))
            if s <= e:
                antig_ranges.append((s, e))
        if not antig_ranges:
            return []
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
    """
    Devuelve por periodo:
      WorkedCost, LostCost, PotentialCost, KPI_COST, LostRate
    """
    if df_cost is None or df_cost.empty:
        return pd.DataFrame()

    windows = build_period_windows(start, end, period).copy()
    if windows.empty:
        return pd.DataFrame()

    # cost dict por cod
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

    # spells (existencia trabajada)
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

    # exits en rango (para LostCost)
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

    # loop por ventana
    rows = []
    for _, w in windows.iterrows():
        ws = pd.Timestamp(w["window_start"]).normalize()
        we = pd.Timestamp(w["window_end"]).normalize()
        ws_day = _day_int(ws)
        we_day = _day_int(we)

        worked_cost = 0.0
        lost_cost = 0.0

        # WorkedCost: sum cost para días trabajados (existencia) elegibles por buckets
        # Recorremos spells con intersección
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

        # LostCost: por cada salida dentro de la ventana, costo de días (exit+1..we)
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
    yellow_min: float = 0.95,  # ratio vs meta
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
# Distribuciones por dimensión (para gráficos “por área/puesto/sexo”)
# - 3 modos: Salidas acumuladas, Existencias snapshot, Existencias promedio (rango)
# =============================================================================
def compute_salidas_by_dim(
    df_events: pd.DataFrame,
    start: pd.Timestamp,
    end: pd.Timestamp,
    dim_col: str,
    antig_sel: List[str],
    edad_sel: List[str],
    unique_personas: bool,
) -> pd.DataFrame:
    _, det = compute_salidas_daily_filtered(
        df_events=df_events,
        start=start,
        end=end,
        antig_sel=antig_sel,
        edad_sel=edad_sel,
        unique_personas_por_dia=True,  # diario da igual, aquí agrupamos
    )
    if det.empty:
        return pd.DataFrame({dim_col: [], "valor": []})

    det[dim_col] = _norm_cat(det[dim_col])
    if unique_personas:
        agg = det.groupby(dim_col)["cod"].nunique().rename("valor").reset_index()
    else:
        agg = det.groupby(dim_col)["cod"].size().rename("valor").reset_index()
    return agg.sort_values("valor", ascending=False).reset_index(drop=True)

def compute_snapshot_by_dim(
    df_events: pd.DataFrame,
    cut: pd.Timestamp,
    dim_col: str,
    antig_sel: List[str],
    edad_sel: List[str],
) -> pd.DataFrame:
    cut = pd.Timestamp(cut).normalize()
    snap = df_events[(df_events["ini"] <= cut) & (df_events["fin_eff"] >= cut)].copy()
    if snap.empty:
        return pd.DataFrame({dim_col: [], "valor": []})

    snap["ref"] = cut
    snap["antig_dias"] = (snap["ref"] - snap["ini"]).dt.days
    snap["Antigüedad"] = bucket_antiguedad(snap["antig_dias"])
    snap["Edad"] = bucket_edad_from_dob(snap["fnac"], snap["ref"])

    if antig_sel:
        snap = snap[snap["Antigüedad"].isin(antig_sel)]
    if edad_sel:
        snap = snap[snap["Edad"].isin(edad_sel)]

    snap[dim_col] = _norm_cat(snap[dim_col])
    agg = snap.groupby(dim_col)["cod"].nunique().rename("valor").reset_index()
    return agg.sort_values("valor", ascending=False).reset_index(drop=True)

def compute_avg_existencias_by_dim(
    df_events: pd.DataFrame,
    start: pd.Timestamp,
    end: pd.Timestamp,
    dim_col: str,
    antig_sel: List[str],
    edad_sel: List[str],
) -> pd.DataFrame:
    start = pd.Timestamp(start).normalize()
    end = pd.Timestamp(end).normalize()
    if start > end:
        return pd.DataFrame({dim_col: [], "valor": []})

    L = int((end - start).days + 1)
    if L <= 0:
        return pd.DataFrame({dim_col: [], "valor": []})

    # spells que intersectan
    d = df_events[(df_events["ini"] <= end) & (df_events["fin_eff"] >= start)].copy()
    if d.empty:
        return pd.DataFrame({dim_col: [], "valor": []})

    d["ini_day"] = d["ini"].apply(_day_int)
    d["fin_day"] = d["fin_eff"].apply(_day_int)
    start_day = _day_int(start)
    end_day = _day_int(end)

    expo_by_cat: Dict[str, int] = {}

    for r in d.itertuples(index=False):
        ini_d = int(r.ini_day)
        fin_d = int(r.fin_day)
        base_s = max(ini_d, start_day)
        base_e = min(fin_d, end_day)
        if base_s > base_e:
            continue

        segs = _segments_for_spell_with_buckets(
            ini_day=ini_d,
            base_s=base_s,
            base_e=base_e,
            dob_ts=r.fnac,
            antig_sel=antig_sel,
            edad_sel=edad_sel,
            start=start,
            end=end,
        )
        if not segs:
            continue

        cat = str(getattr(r, dim_col))
        cat = _norm_cat(pd.Series([cat])).iloc[0]

        expo = 0
        for (ss, ee) in segs:
            expo += (ee - ss + 1)

        expo_by_cat[cat] = expo_by_cat.get(cat, 0) + expo

    if not expo_by_cat:
        return pd.DataFrame({dim_col: [], "valor": []})

    out = pd.DataFrame({dim_col: list(expo_by_cat.keys()), "expo": list(expo_by_cat.values())})
    out["valor"] = out["expo"].astype(float) / float(L)
    out = out.drop(columns=["expo"]).sort_values("valor", ascending=False).reset_index(drop=True)
    return out


# =============================================================================
# Contingencia
# =============================================================================
DIMENSIONS = {
    "Área General": "area_gen",
    "Área": "area",
    "Cargo Actual": "cargo",
    "Clasificación": "clas",
    "Sexo": "sexo",
    "Trabajadora Social": "ts",
    "Empresa": "emp",
    "Nacionalidad": "nac",
    "Lugar Registro": "lug",
    "Región Registro": "reg",
    "Antigüedad (bucket)": "Antigüedad",
    "Edad (bucket)": "Edad",
}

def contingency_tables(
    df: pd.DataFrame,
    row_dim: str,
    col_dim: str,
    top_rows: int = 40,
    top_cols: int = 40,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    r = DIMENSIONS[row_dim]
    c = DIMENSIONS[col_dim]

    rr = _norm_cat(df[r])
    cc = _norm_cat(df[c])

    counts = pd.crosstab(rr, cc, dropna=False)
    counts = _reduce_crosstab(counts, top_rows=top_rows, top_cols=top_cols)

    denom = counts.sum(axis=1).replace(0, np.nan)
    pct = counts.div(denom, axis=0) * 100

    counts.index = counts.index.map(str)
    counts.columns = counts.columns.map(str)
    pct.index = pct.index.map(str)
    pct.columns = pct.columns.map(str)

    return counts, pct.round(2)


# =============================================================================
# Plot helpers
# =============================================================================
def apply_bar_labels(fig, show_labels: bool, fmt: str = ".0f"):
    if show_labels:
        fig.update_traces(texttemplate=f"%{{y:{fmt}}}", textposition="outside", cliponaxis=False)
    return fig

def apply_line_labels(fig, show_labels: bool, fmt: str = ".2f"):
    if show_labels:
        fig.update_traces(mode="lines+markers+text", texttemplate=f"%{{y:{fmt}}}", textposition="top center")
    return fig

def nice_xaxis(fig):
    fig.update_xaxes(type="category", automargin=True)
    fig.update_layout(margin=dict(b=80))
    return fig


# =============================================================================
# Layout general: Main + panel derecho (filtros)
# =============================================================================
main_col, filter_col = st.columns([4.3, 1.7], gap="large")

# =============================================================================
# Panel derecho
# =============================================================================
with filter_col:
    st.subheader("Panel de control")

    tab_p, tab_f, tab_o, tab_c = st.tabs(["Datos & Periodo", "Filtros", "Opciones", "Costo KPI"])

    # -------------------------
    # Datos & Periodo
    # -------------------------
    with tab_p:
        st.markdown("**Carga de datos**")

        uploaded = st.file_uploader("Sube Excel/CSV (Historia Personal)", type=["xlsx", "xls", "csv"], key="uploader_hist")
        path = st.text_input("O ruta local (opcional)", value="", key="path_hist")

        df_raw = None
        df_cost_raw = None
        sheet_hist = None
        sheet_cost = None

        if uploaded is None and not path.strip():
            st.info("Carga un archivo para iniciar.")
            st.stop()

        try:
            if uploaded is not None:
                if uploaded.name.lower().endswith(".csv"):
                    df_raw = read_csv_any(uploaded)
                else:
                    xls = pd.ExcelFile(uploaded)
                    sheet_hist = st.selectbox("Hoja (Historia Personal)", options=xls.sheet_names, index=0, key="sheet_hist_upload")
                    df_raw = read_excel_strict_hist(uploaded, sheet_hist)

                    sheet_cost = find_cost_sheet_name(xls)
                    if sheet_cost:
                        with st.expander("Costo Nominal detectado (opcional)", expanded=False):
                            st.caption(f"Hoja detectada: {sheet_cost}")
                            use_cost = st.checkbox("Usar Costo Nominal de este mismo Excel", value=True, key="use_cost_same")
                            if use_cost:
                                df_cost_raw = read_excel_any(uploaded, sheet_cost)
                            else:
                                df_cost_raw = None
                    else:
                        with st.expander("Costo Nominal (opcional)", expanded=False):
                            st.caption("No detecté hoja 'Costo Nominal' en este Excel.")
                            up_cost = st.file_uploader("Sube Excel (Costo Nominal)", type=["xlsx", "xls"], key="uploader_cost")
                            if up_cost is not None:
                                x2 = pd.ExcelFile(up_cost)
                                sheet_cost = st.selectbox("Hoja (Costo Nominal)", options=x2.sheet_names, index=0, key="sheet_cost_upload")
                                df_cost_raw = read_excel_any(up_cost, sheet_cost)

            else:
                # path local
                p = path.strip()
                if not os.path.exists(p):
                    st.error("La ruta no existe.")
                    st.stop()
                if p.lower().endswith(".csv"):
                    df_raw = read_csv_any(p)
                else:
                    xls = pd.ExcelFile(p)
                    sheet_hist = st.selectbox("Hoja (Historia Personal)", options=xls.sheet_names, index=0, key="sheet_hist_path")
                    df_raw = read_excel_strict_hist(p, sheet_hist)

                    sheet_cost = find_cost_sheet_name(xls)
                    if sheet_cost:
                        with st.expander("Costo Nominal detectado (opcional)", expanded=False):
                            use_cost = st.checkbox("Usar Costo Nominal de este mismo Excel", value=True, key="use_cost_same_path")
                            if use_cost:
                                df_cost_raw = read_excel_any(p, sheet_cost)
                            else:
                                df_cost_raw = None
                    else:
                        with st.expander("Costo Nominal (opcional)", expanded=False):
                            st.caption("No detecté hoja 'Costo Nominal' en este Excel.")
                            cost_path = st.text_input("Ruta local Costo Nominal (opcional)", value="", key="path_cost")
                            if cost_path.strip() and os.path.exists(cost_path.strip()):
                                x2 = pd.ExcelFile(cost_path.strip())
                                sheet_cost = st.selectbox("Hoja (Costo Nominal)", options=x2.sheet_names, index=0, key="sheet_cost_path")
                                df_cost_raw = read_excel_any(cost_path.strip(), sheet_cost)

        except Exception as e:
            st.error(f"No se pudo leer el archivo: {e}")
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

        st.markdown("**Rango de análisis (Vista Historia)**")
        min_date = df_intervals_all["ini"].min()
        max_date = df_intervals_all["fin_eff"].max()
        default_end = min(today_dt(), max_date) if pd.notna(max_date) else today_dt()
        default_start = max(min_date, default_end - pd.Timedelta(days=180)) if pd.notna(min_date) else (default_end - pd.Timedelta(days=180))

        preset = st.selectbox(
            "Atajo de rango",
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
            "Inicio / Fin",
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

        period_label = st.selectbox("Agrupar por", options=["Día", "Semana", "Mes", "Año"], index=1, key="period_group")
        period = {"Día": "D", "Semana": "W", "Mes": "M", "Año": "Y"}[period_label]

        snap_date = st.slider(
            "Snapshot existencias (día)",
            min_value=start_dt.date(),
            max_value=end_dt.date(),
            value=end_dt.date(),
            key="snap_date",
        )
        snap_dt = pd.Timestamp(snap_date)

        st.markdown("**Corte para Vista Actual (hoy)**")
        cut_today = min(today_dt(), max_date) if pd.notna(max_date) else today_dt()
        st.write(f"Hoy (corte): **{cut_today.date()}**")

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
    # Filtros (multi-select)
    # -------------------------
    with tab_f:
        g = st.session_state.get("__globals__")
        if not g:
            st.stop()
        df0 = g["df0"]

        st.markdown("**Filtros globales (multi-select)**")
        st.caption("Deja vacío = no filtra (equivale a TODOS).")

        def opts(df: pd.DataFrame, col: str) -> List[str]:
            v = df[col].dropna().astype(str).str.strip()
            v = v[v != ""].unique().tolist()
            return sorted(v)

        if st.button("Limpiar filtros", use_container_width=True, key="btn_clear_filters"):
            for k in [
                "f_sexo", "f_area_gen", "f_area", "f_cargo", "f_clas", "f_ts", "f_emp", "f_nac", "f_lug", "f_reg",
                "f_antig", "f_edad",
            ]:
                st.session_state[k] = []
            st.rerun()

        area_gen_pick = st.multiselect("Área General", opts(df0, "area_gen"), default=st.session_state.get("f_area_gen", []), key="f_area_gen")

        if area_gen_pick:
            df_area = df0[df0["area_gen"].isin(area_gen_pick)]
            area_opts = opts(df_area, "area")
        else:
            area_opts = opts(df0, "area")

        fs = FilterState(
            sexo=st.multiselect("Sexo", opts(df0, "sexo"), default=st.session_state.get("f_sexo", []), key="f_sexo"),
            area_gen=area_gen_pick,
            area=st.multiselect("Área (nombre)", area_opts, default=st.session_state.get("f_area", []), key="f_area"),
            cargo=st.multiselect("Cargo", opts(df0, "cargo"), default=st.session_state.get("f_cargo", []), key="f_cargo"),
            clas=st.multiselect("Clasificación", opts(df0, "clas"), default=st.session_state.get("f_clas", []), key="f_clas"),
            ts=st.multiselect("Trabajadora Social", opts(df0, "ts"), default=st.session_state.get("f_ts", []), key="f_ts"),
            emp=st.multiselect("Empresa", opts(df0, "emp"), default=st.session_state.get("f_emp", []), key="f_emp"),
            nac=st.multiselect("Nacionalidad", opts(df0, "nac"), default=st.session_state.get("f_nac", []), key="f_nac"),
            lug=st.multiselect("Lugar Registro", opts(df0, "lug"), default=st.session_state.get("f_lug", []), key="f_lug"),
            reg=st.multiselect("Región Registro", opts(df0, "reg"), default=st.session_state.get("f_reg", []), key="f_reg"),
            antig=st.multiselect(
                "Antigüedad (bucket)",
                ["< 30 días", "30 - 90 días", "91 - 180 días", "181 - 360 días", "> 360 días", MISSING_LABEL],
                default=st.session_state.get("f_antig", []),
                key="f_antig",
            ),
            edad=st.multiselect(
                "Edad (bucket)",
                ["< 24 años", "24 - 30 años", "31 - 37 años", "38 - 42 años", "43 - 49 años", "50 - 56 años", "> 56 años", MISSING_LABEL],
                default=st.session_state.get("f_edad", []),
                key="f_edad",
            ),
        )
        st.session_state["__fs__"] = fs

    # -------------------------
    # Opciones
    # -------------------------
    with tab_o:
        unique_personas_por_dia = st.checkbox("Salidas: contar personas únicas por día", value=True, key="opt_unique_day")
        show_labels = st.checkbox("Mostrar etiquetas numéricas", value=False, key="opt_show_labels")

        st.markdown("**Horizonte H (PDE equivalente)**")
        h_choice = st.selectbox("Horizonte", options=["30 días", "90 días", "180 días", "Otro…"], index=0, key="opt_h_choice")
        if h_choice == "30 días":
            horizon_days = 30
        elif h_choice == "90 días":
            horizon_days = 90
        elif h_choice == "180 días":
            horizon_days = 180
        else:
            horizon_days = int(st.number_input("H (días)", min_value=7, max_value=365, value=30, step=1, key="opt_h_custom"))

        st.markdown("**Semáforo KPI_PDE (ratio vs meta)**")
        green_max = float(st.number_input("Verde si ≤", min_value=0.70, max_value=1.10, value=0.95, step=0.01, key="opt_green"))
        yellow_max = float(st.number_input("Amarillo si ≤", min_value=0.80, max_value=1.30, value=1.05, step=0.01, key="opt_yellow"))

        st.markdown("**Contingencia**")
        top_rows = int(st.number_input("Máx. filas", min_value=10, max_value=200, value=40, step=5, key="opt_top_rows"))
        top_cols = int(st.number_input("Máx. columnas", min_value=10, max_value=200, value=40, step=5, key="opt_top_cols"))

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
    # Costo KPI
    # -------------------------
    with tab_c:
        st.caption("Si no cargaste 'Costo Nominal', esta parte se desactiva sola.")
        yellow_min_cost = float(st.number_input("AMARILLO si ratio vs meta ≥", min_value=0.50, max_value=0.99, value=0.95, step=0.01, key="opt_cost_yellow"))
        st.session_state["__cost_opts__"] = {"yellow_min_cost": yellow_min_cost}


# =============================================================================
# Cálculos globales (Historia + Actual)
# =============================================================================
g = st.session_state.get("__globals__")
fs = st.session_state.get("__fs__")
opts = st.session_state.get("__opts__")

if not g or not fs or not opts:
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
top_rows = int(opts["top_rows"])
top_cols = int(opts["top_cols"])
yellow_min_cost = float(st.session_state.get("__cost_opts__", {}).get("yellow_min_cost", 0.95))

with st.spinner("Calculando métricas (Historia y Actual)..."):
    # aplica filtros categóricos
    df0_f = apply_categorical_filters(df0, fs)

    # intervalos para existencias (global del filtro)
    df_intervals_f = merge_intervals_per_person(df0_f) if not df0_f.empty else df0_f.copy()

    # --- Historia (rango seleccionado)
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

    # Baseline (empresa completa) mismo rango y mismos buckets
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

    # PDE por periodo + KPI PDE
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

    # Snapshot (historia) en snap_dt
    df_snap_hist = df0_f[(df0_f["ini"] <= snap_dt) & (df0_f["fin_eff"] >= snap_dt)].copy()
    if not df_snap_hist.empty:
        df_snap_hist["ref"] = snap_dt
        df_snap_hist["antig_dias"] = (df_snap_hist["ref"] - df_snap_hist["ini"]).dt.days
        df_snap_hist["Antigüedad"] = bucket_antiguedad(df_snap_hist["antig_dias"])
        df_snap_hist["Edad"] = bucket_edad_from_dob(df_snap_hist["fnac"], df_snap_hist["ref"])
        if fs.antig:
            df_snap_hist = df_snap_hist[df_snap_hist["Antigüedad"].isin(fs.antig)]
        if fs.edad:
            df_snap_hist = df_snap_hist[df_snap_hist["Edad"].isin(fs.edad)]

    # KPI COSTO (si existe costo nominal)
    df_cost_g = compute_cost_period_metrics(
        df_events=df0_f,
        df_cost=df_cost if df_cost is not None else pd.DataFrame(),
        start=start_dt,
        end=end_dt,
        period=period,
        antig_sel=fs.antig,
        edad_sel=fs.edad,
        unique_salidas_por_dia=unique_personas_por_dia,
    ) if (df_cost is not None and not df_cost.empty) else pd.DataFrame()

    df_cost_b = compute_cost_period_metrics(
        df_events=df0,
        df_cost=df_cost if df_cost is not None else pd.DataFrame(),
        start=start_dt,
        end=end_dt,
        period=period,
        antig_sel=fs.antig,
        edad_sel=fs.edad,
        unique_salidas_por_dia=unique_personas_por_dia,
    ) if (df_cost is not None and not df_cost.empty) else pd.DataFrame()

    df_kpi_cost = compute_cost_kpis_vs_meta(
        df_cost_g=df_cost_g,
        df_cost_b=df_cost_b,
        period=period,
        yellow_min=yellow_min_cost,
    ) if (df_cost_g is not None and not df_cost_g.empty) else pd.DataFrame()

    # --- Vista Actual (corte hoy)
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

    # resumen por persona (para análisis Actual)
    # - N° Ingresos: cantidad de spells (ini)
    # - Antigüedad Último Ingreso
    # - Antigüedad Acumulada (suma días trabajados)
    df_all = df0_f.copy()
    df_all["ini_n"] = df_all["ini"].fillna(pd.NaT)
    df_all["fin_n"] = df_all["fin_eff"].fillna(cut_today)
    df_all["fin_n"] = df_all["fin_n"].clip(upper=cut_today)

    # días trabajados por spell hasta hoy
    df_all["days_spell"] = (df_all["fin_n"] - df_all["ini_n"]).dt.days + 1
    df_all["days_spell"] = df_all["days_spell"].clip(lower=0).fillna(0).astype(int)

    per = df_all.groupby("cod", as_index=False).agg(
        N_Ingresos=("ini_n", "count"),
        Ultimo_Ingreso=("ini_n", "max"),
        Antig_Acumulada_Dias=("days_spell", "sum"),
        r_pct=("r_pct", "last"),
        area_gen=("area_gen", "last"),
        area=("area", "last"),
        cargo=("cargo", "last"),
        sexo=("sexo", "last"),
        emp=("emp", "last"),
        ts=("ts", "last"),
        clas=("clas", "last"),
        nac=("nac", "last"),
        reg=("reg", "last"),
    )
    per["Antig_UltimoIngreso_Dias"] = (cut_today - per["Ultimo_Ingreso"]).dt.days
    per["Antig_UltimoIngreso_Dias"] = per["Antig_UltimoIngreso_Dias"].fillna(np.nan)
    per["Antig_Acumulada_Dias"] = per["Antig_Acumulada_Dias"].astype(int)

    # solo activos hoy (para tabla principal “Actual”)
    activos_hoy = set(df_now["cod"].unique().tolist()) if df_now is not None and not df_now.empty else set()
    per_act = per[per["cod"].isin(activos_hoy)].copy()


# =============================================================================
# UI principal: 2 VISTAS (Historia vs Actual)
# =============================================================================
with main_col:
    vista = st.tabs(["📈 Historia (Salidas & Existencias)", "🟢 Actual (Existencias hoy)"])

    # =====================================================================
    # VISTA 1: HISTORIA
    # =====================================================================
    with vista[0]:
        # KPIs arriba
        sal_total = int(df_daily_g["Salidas"].sum()) if not df_daily_g.empty else 0
        exist_prom = float(df_daily_g["Existencias"].mean()) if not df_daily_g.empty else 0.0
        exist_snap = int(df_snap_hist["cod"].nunique()) if (df_snap_hist is not None and not df_snap_hist.empty) else 0

        pde_last = np.nan
        kpi_pde_last = np.nan
        sem_pde = "-"
        if not df_kpi_pde.empty and df_kpi_pde["PDEH_g"].dropna().shape[0] > 0:
            last = df_kpi_pde.dropna(subset=["PDEH_g"]).iloc[-1]
            pde_last = float(last["PDEH_g"])
            kpi_pde_last = float(last["KPI_PDE"]) if pd.notna(last.get("KPI_PDE", np.nan)) else np.nan
            sem_pde = str(last.get("Semaforo", "-"))

        kpi_cost_last = np.nan
        sem_cost = "-"
        if df_kpi_cost is not None and not df_kpi_cost.empty and df_kpi_cost["KPI_COST"].dropna().shape[0] > 0:
            lastc = df_kpi_cost.dropna(subset=["KPI_COST"]).iloc[-1]
            kpi_cost_last = float(lastc["KPI_COST"])
            sem_cost = str(lastc.get("Semaforo_COST", "-"))

        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("Salidas (rango)", f"{sal_total:,}".replace(",", "."))
        k2.metric("Existencias prom (rango)", f"{exist_prom:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        k3.metric(f"Existencias snapshot ({snap_dt.date()})", f"{exist_snap:,}".replace(",", "."))
        k4.metric(f"PDE{horizon_days} (último)", "-" if np.isnan(pde_last) else f"{pde_last:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
        k5.metric("KPI_PDE (último)", "-" if np.isnan(kpi_pde_last) else f"{kpi_pde_last:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
        k6.metric("KPI_COST (último)", "-" if np.isnan(kpi_cost_last) else f"{kpi_cost_last:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))

        # Estructura tipo panel (sin copiar colores): 2 columnas de navegación
        left, right = st.columns([2.2, 1.2], gap="large")

        with left:
            st.subheader("Evolución (rango seleccionado)")

            # Serie: Existencias vs Salidas (misma figura, doble eje)
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=df_daily_g["Día"],
                y=df_daily_g["Existencias"],
                name="Existencias",
                mode="lines",
                yaxis="y1",
            ))
            fig.add_trace(go.Bar(
                x=df_daily_g["Día"],
                y=df_daily_g["Salidas"],
                name="Salidas",
                yaxis="y2",
                opacity=0.6,
            ))
            fig.update_layout(
                title="Existencias (línea) vs Salidas (barras)",
                xaxis=dict(title="Día", automargin=True),
                yaxis=dict(title="Existencias", side="left"),
                yaxis2=dict(title="Salidas", overlaying="y", side="right", showgrid=False),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
                margin=dict(b=80),
            )
            st.plotly_chart(fig, use_container_width=True)

            st.markdown("---")
            st.subheader("Distribuciones (elige DIM y MÉTRICA por gráfico)")

            dim_opts = [
                ("Área General", "area_gen"),
                ("Área", "area"),
                ("Cargo", "cargo"),
                ("Clasificación", "clas"),
                ("Empresa", "emp"),
                ("TS", "ts"),
                ("Nacionalidad", "nac"),
                ("Región", "reg"),
            ]
            dim_labels = [x[0] for x in dim_opts]
            dim_map = {lab: col for lab, col in dim_opts}

            metric_opts = [
                "Salidas acumuladas (rango)",
                f"Existencias snapshot ({snap_dt.date()})",
                "Existencias promedio (rango)",
            ]

            cA, cB, cC = st.columns([1.2, 1.2, 1.0], gap="large")

            # --------- BAR 1
            with cA:
                dim1_lab = st.selectbox("Gráfico 1 (barras) - Dimensión", dim_labels, index=0, key="g1_dim")
                met1 = st.selectbox("Gráfico 1 - Métrica", metric_opts, index=1, key="g1_met")

                dim1 = dim_map[dim1_lab]
                if met1.startswith("Salidas"):
                    d1 = compute_salidas_by_dim(df0_f, start_dt, end_dt, dim1, fs.antig, fs.edad, unique_personas=True)
                elif met1.startswith("Existencias snapshot"):
                    d1 = compute_snapshot_by_dim(df0_f, snap_dt, dim1, fs.antig, fs.edad)
                else:
                    d1 = compute_avg_existencias_by_dim(df0_f, start_dt, end_dt, dim1, fs.antig, fs.edad)

                d1 = d1.head(25)
                fig1 = px.bar(d1, x="valor", y=dim1, orientation="h", title=f"{met1} por {dim1_lab}")
                fig1.update_layout(yaxis_title="")
                st.plotly_chart(apply_bar_labels(fig1, show_labels, ".2f" if "promedio" in met1 else ".0f"), use_container_width=True)

            # --------- BAR 2
            with cB:
                dim2_lab = st.selectbox("Gráfico 2 (barras) - Dimensión", dim_labels, index=2, key="g2_dim")
                met2 = st.selectbox("Gráfico 2 - Métrica", metric_opts, index=0, key="g2_met")

                dim2 = dim_map[dim2_lab]
                if met2.startswith("Salidas"):
                    d2 = compute_salidas_by_dim(df0_f, start_dt, end_dt, dim2, fs.antig, fs.edad, unique_personas=True)
                elif met2.startswith("Existencias snapshot"):
                    d2 = compute_snapshot_by_dim(df0_f, snap_dt, dim2, fs.antig, fs.edad)
                else:
                    d2 = compute_avg_existencias_by_dim(df0_f, start_dt, end_dt, dim2, fs.antig, fs.edad)

                d2 = d2.head(25)
                fig2 = px.bar(d2, x="valor", y=dim2, orientation="h", title=f"{met2} por {dim2_lab}")
                fig2.update_layout(yaxis_title="")
                st.plotly_chart(apply_bar_labels(fig2, show_labels, ".2f" if "promedio" in met2 else ".0f"), use_container_width=True)

            # --------- PIE
            with cC:
                dim3_lab = st.selectbox("Gráfico 3 (pie) - Dimensión", ["Sexo"] + dim_labels, index=0, key="g3_dim")
                met3 = st.selectbox("Gráfico 3 - Métrica", metric_opts, index=1, key="g3_met")

                dim3 = "sexo" if dim3_lab == "Sexo" else dim_map[dim3_lab]
                if met3.startswith("Salidas"):
                    d3 = compute_salidas_by_dim(df0_f, start_dt, end_dt, dim3, fs.antig, fs.edad, unique_personas=True)
                elif met3.startswith("Existencias snapshot"):
                    d3 = compute_snapshot_by_dim(df0_f, snap_dt, dim3, fs.antig, fs.edad)
                else:
                    d3 = compute_avg_existencias_by_dim(df0_f, start_dt, end_dt, dim3, fs.antig, fs.edad)

                d3 = d3.head(12)
                if d3.empty:
                    st.info("Sin datos para el pie con estos filtros.")
                else:
                    fig3 = px.pie(d3, names=dim3, values="valor", title=f"{met3} por {dim3_lab}")
                    st.plotly_chart(fig3, use_container_width=True)

            st.markdown("---")
            st.subheader("Tabla cruzada (contingencia)")

            # Para salidas (detalle en rango) o snapshot (existencias)
            base_choice = st.radio("Base de la contingencia", ["Salidas (rango)", f"Existencias (snapshot {snap_dt.date()})"], horizontal=True, key="ct_base")
            if base_choice.startswith("Salidas"):
                df_ctx = df_sal_det.copy()
                if not df_ctx.empty:
                    df_ctx = df_ctx.copy()
                    # ya viene con Antigüedad/Edad (en compute_salidas_daily_filtered)
            else:
                df_ctx = df_snap_hist.copy()

            if df_ctx is None or df_ctx.empty:
                st.warning("No hay registros para contingencia con los filtros actuales.")
            else:
                # asegura buckets si el contexto es snapshot
                if base_choice.startswith("Existencias") and "Antigüedad" not in df_ctx.columns:
                    df_ctx = df_ctx.copy()
                    df_ctx["ref"] = snap_dt
                    df_ctx["antig_dias"] = (df_ctx["ref"] - df_ctx["ini"]).dt.days
                    df_ctx["Antigüedad"] = bucket_antiguedad(df_ctx["antig_dias"])
                    df_ctx["Edad"] = bucket_edad_from_dob(df_ctx["fnac"], df_ctx["ref"])

                row_dim = st.selectbox("Filas", list(DIMENSIONS.keys()), index=list(DIMENSIONS.keys()).index("Área General"), key="ct_row")
                col_dim = st.selectbox("Columnas", list(DIMENSIONS.keys()), index=list(DIMENSIONS.keys()).index("Clasificación"), key="ct_col")

                counts, pct = contingency_tables(df_ctx, row_dim, col_dim, top_rows=top_rows, top_cols=top_cols)

                t1, t2 = st.columns(2)
                with t1:
                    st.markdown("**Conteos**")
                    st.dataframe(_safe_table_for_streamlit(counts.reset_index().rename(columns={"index": row_dim})), use_container_width=True, height=420)
                with t2:
                    st.markdown("**% por fila**")
                    st.dataframe(_safe_table_for_streamlit(pct.reset_index().rename(columns={"index": row_dim})), use_container_width=True, height=420)

        with right:
            st.subheader("KPIs (por periodo)")

            # KPI PDE y Meta
            if df_kpi_pde.empty or df_kpi_pde["KPI_PDE"].dropna().empty:
                st.info("No hay suficiente data para KPI_PDE con este rango/filtros.")
            else:
                figk = px.line(df_kpi_pde, x="Periodo", y=["KPI_PDE"], title="KPI_PDE (PDE_H / Meta)")
                st.plotly_chart(nice_xaxis(figk), use_container_width=True)

                figp = px.line(df_kpi_pde, x="Periodo", y=["PDEH_g", "Meta"], title=f"PDE{horizon_days} vs Meta (baseline)")
                st.plotly_chart(nice_xaxis(figp), use_container_width=True)

            st.markdown("---")
            st.subheader("KPI COSTO (nuevo)")

            if df_kpi_cost is None or df_kpi_cost.empty:
                st.info("Costo KPI desactivado (no hay hoja Costo Nominal o no pudo leerse).")
            else:
                figc1 = px.line(df_kpi_cost, x="Periodo", y=["KPI_COST", "Meta_COST"], title="KPI_COST vs Meta_COST")
                st.plotly_chart(nice_xaxis(figc1), use_container_width=True)

                figc2 = px.line(df_kpi_cost, x="Periodo", y="Indice_vs_Meta_COST", title="Índice vs Meta (KPI_COST / Meta_COST)")
                st.plotly_chart(nice_xaxis(figc2), use_container_width=True)

                st.dataframe(
                    _safe_table_for_streamlit(
                        df_kpi_cost[["Periodo", "KPI_COST", "Meta_COST", "Indice_vs_Meta_COST", "Semaforo_COST", "WorkedCost", "LostCost", "PotentialCost"]].tail(60)
                    ),
                    use_container_width=True,
                    height=320,
                )

            st.markdown("---")
            st.subheader("Descargas (Historia)")
            buf_xlsx = io.BytesIO()
            with pd.ExcelWriter(buf_xlsx, engine="openpyxl") as writer:
                df_daily_g.to_excel(writer, index=False, sheet_name="Diario_Grupo")
                df_period_g.to_excel(writer, index=False, sheet_name="Periodo_PDE_Grupo")
                df_kpi_pde.to_excel(writer, index=False, sheet_name="KPI_PDE")
                df_sal_det.to_excel(writer, index=False, sheet_name="Salidas_Detalle")
                df_snap_hist.to_excel(writer, index=False, sheet_name="Snapshot")
                if df_kpi_cost is not None and not df_kpi_cost.empty:
                    df_kpi_cost.to_excel(writer, index=False, sheet_name="KPI_COST")
            st.download_button(
                "Descargar Excel (Historia)",
                data=buf_xlsx.getvalue(),
                file_name="rrhh_historia.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_historia",
            )

    # =====================================================================
    # VISTA 2: ACTUAL
    # =====================================================================
    with vista[1]:
        st.subheader(f"Existencias actuales (corte {cut_today.date()}) + acumulados hasta hoy")

        exist_hoy = int(df_now["cod"].nunique()) if df_now is not None and not df_now.empty else 0

        # acumulados: por defecto desde inicio de año hasta hoy (editable)
        colA, colB = st.columns([1.2, 2.8], gap="large")
        with colA:
            default_acc_start = date(cut_today.year, 1, 1)
            acc_start = st.date_input("Acumulado desde", value=default_acc_start, key="acc_start")
            acc_start_dt = pd.Timestamp(acc_start)
            acc_end_dt = cut_today

            unique_ing = st.checkbox("Ingresos: contar personas únicas (no registros)", value=True, key="acc_unique_ing")
            unique_sal = st.checkbox("Salidas: contar personas únicas por día (para costo/pde)", value=True, key="acc_unique_sal")

        # acumulados
        ingresos_acc = compute_ingresos_total(df0_f, acc_start_dt, acc_end_dt, unique_personas_en_rango=unique_ing)
        salidas_acc = int(
            compute_salidas_by_dim(df0_f, acc_start_dt, acc_end_dt, "emp", fs.antig, fs.edad, unique_personas=True)["valor"].sum()
        ) if not df0_f.empty else 0

        # indicadores extra: “rotación” simple acumulada (proxy)
        rot_simple = (salidas_acc / max(exist_hoy, 1)) * 100.0 if exist_hoy > 0 else np.nan

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Existencias hoy", f"{exist_hoy:,}".replace(",", "."))
        k2.metric(f"Ingresos acumulados ({acc_start_dt.date()}→hoy)", f"{ingresos_acc:,}".replace(",", "."))
        k3.metric(f"Salidas acumuladas ({acc_start_dt.date()}→hoy)", f"{salidas_acc:,}".replace(",", "."))
        k4.metric("Rotación simple (%)", "-" if np.isnan(rot_simple) else f"{rot_simple:,.1f}%".replace(",", "X").replace(".", ",").replace("X", "."))

        with colB:
            st.markdown("### Distribución actual (snapshot hoy)")
            if df_now is None or df_now.empty:
                st.warning("No hay existencias activas hoy con los filtros actuales.")
            else:
                c1, c2, c3 = st.columns(3)
                with c1:
                    d = compute_snapshot_by_dim(df0_f, cut_today, "area_gen", fs.antig, fs.edad).head(20)
                    fig = px.bar(d, x="valor", y="area_gen", orientation="h", title="Existencias hoy por Área General")
                    fig.update_layout(yaxis_title="")
                    st.plotly_chart(fig, use_container_width=True)
                with c2:
                    d = compute_snapshot_by_dim(df0_f, cut_today, "cargo", fs.antig, fs.edad).head(20)
                    fig = px.bar(d, x="valor", y="cargo", orientation="h", title="Existencias hoy por Cargo")
                    fig.update_layout(yaxis_title="")
                    st.plotly_chart(fig, use_container_width=True)
                with c3:
                    d = compute_snapshot_by_dim(df0_f, cut_today, "sexo", fs.antig, fs.edad).head(10)
                    if d.empty:
                        st.info("Sin datos para sexo.")
                    else:
                        fig = px.pie(d, names="sexo", values="valor", title="Existencias hoy por Sexo")
                        st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.subheader("Análisis con variables nuevas (por persona)")
        st.caption("Antigüedad Último Ingreso, Antigüedad Acumulada, N° Ingresos, %R (si existe).")

        if per_act is None or per_act.empty:
            st.info("No hay tabla de personas activas hoy con estos filtros.")
        else:
            a1, a2 = st.columns([1.4, 1.6], gap="large")
            with a1:
                st.markdown("**Resumen**")
                st.dataframe(
                    _safe_table_for_streamlit(
                        per_act[[
                            "cod", "N_Ingresos", "Ultimo_Ingreso",
                            "Antig_UltimoIngreso_Dias", "Antig_Acumulada_Dias",
                            "r_pct", "area_gen", "area", "cargo", "sexo"
                        ]].sort_values(["Antig_Acumulada_Dias"], ascending=False)
                    ),
                    use_container_width=True,
                    height=420,
                )

            with a2:
                sel = st.selectbox(
                    "Gráfico (distribución)",
                    options=[
                        "Antigüedad Último Ingreso (días)",
                        "Antigüedad Acumulada (días)",
                        "N° Ingresos",
                        "%R (factor)",
                    ],
                    index=0,
                    key="act_dist_pick",
                )

                if sel == "Antigüedad Último Ingreso (días)":
                    fig = px.histogram(per_act, x="Antig_UltimoIngreso_Dias", nbins=30, title="Distribución: Antigüedad Último Ingreso (días)")
                    st.plotly_chart(fig, use_container_width=True)
                elif sel == "Antigüedad Acumulada (días)":
                    fig = px.histogram(per_act, x="Antig_Acumulada_Dias", nbins=30, title="Distribución: Antigüedad Acumulada (días)")
                    st.plotly_chart(fig, use_container_width=True)
                elif sel == "N° Ingresos":
                    fig = px.histogram(per_act, x="N_Ingresos", nbins=20, title="Distribución: N° Ingresos")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    fig = px.histogram(per_act, x="r_pct", nbins=25, title="Distribución: %R (factor)")
                    st.plotly_chart(fig, use_container_width=True)

                st.markdown("**Relación (scatter)**")
                fig2 = px.scatter(
                    per_act,
                    x="Antig_UltimoIngreso_Dias",
                    y="Antig_Acumulada_Dias",
                    size="N_Ingresos",
                    hover_data=["cod", "area_gen", "cargo", "r_pct"],
                    title="Antig. último ingreso vs Antig. acumulada (tamaño = N° ingresos)",
                )
                st.plotly_chart(fig2, use_container_width=True)

        st.markdown("---")
        st.subheader("Descarga (Actual)")
        buf_now = io.BytesIO()
        with pd.ExcelWriter(buf_now, engine="openpyxl") as writer:
            if df_now is not None:
                df_now.to_excel(writer, index=False, sheet_name="Existencias_Hoy")
            if per_act is not None:
                per_act.to_excel(writer, index=False, sheet_name="Personas_Activas_Analisis")
        st.download_button(
            "Descargar Excel (Actual)",
            data=buf_now.getvalue(),
            file_name="rrhh_actual.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_actual",
        )


st.divider()
st.caption(
    "Listo: filtros multi-select, 2 vistas (Historia/Actual), panel tipo dashboard para navegar fácil, "
    "series de tiempo (Existencias vs Salidas + KPI_PDE + KPI_COST con Meta), distribuciones configurables por gráfico, "
    "contingencia configurable, y análisis en vista Actual con Antigüedades / N° ingresos / %R."
)
