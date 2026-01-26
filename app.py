from __future__ import annotations

import io
import os
from dataclasses import dataclass
from datetime import date, timedelta
from typing import List, Tuple, Optional, Dict

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# =========================
# Config UI
# =========================
st.set_page_config(page_title="PDE por Exposición (persona-días) / Existencias / Salidas", layout="wide")
st.title("Dashboard: PDE por Exposición (persona-días) / Existencias / Salidas")


# =========================
# Columnas esperadas
# =========================
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


# =========================
# TABLAS DE REFERENCIA
# =========================
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


# =========================
# Helpers generales
# =========================
def _to_datetime(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce").dt.normalize()

def today_dt() -> pd.Timestamp:
    return pd.Timestamp(date.today())

def excel_weeknum_return_type_1(d: pd.Series) -> pd.Series:
    # Excel NUM.DE.SEMANA(fecha;1) ~ strftime %U + 1 (domingo->sábado)
    return d.dt.strftime("%U").astype(int) + 1

def week_end_sun_to_sat(d: pd.Series) -> pd.Series:
    # fin de semana (sábado) para semanas tipo Excel (domingo->sábado)
    wd = d.dt.weekday  # Mon=0..Sun=6
    days_since_sun = (wd + 1) % 7  # Sun->0, Mon->1, ...
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


# =========================
# Mapping de Área y Clasificación
# =========================
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


# =========================
# Preparación de datos
# =========================
KEEP_INTERNAL = [
    "cod", "ini", "fin", "fin_eff", "fnac",
    "clas_raw", "clas",
    "sexo", "ts", "emp",
    "area_raw", "area", "area_gen",
    "cargo", "nac", "lug", "reg",
]

@st.cache_data(show_spinner=False)
def validate_and_prepare(df_raw: pd.DataFrame) -> pd.DataFrame:
    missing = [c for c in REQUIRED_COLS if c not in df_raw.columns]
    if missing:
        raise ValueError("Faltan columnas requeridas:\n- " + "\n- ".join(missing))

    df = df_raw[REQUIRED_COLS].copy()
    out = df.rename(columns=COL_MAP)

    out["ini"] = _to_datetime(out["ini"])
    out["fin"] = _to_datetime(out["fin"])
    out["fnac"] = _to_datetime(out["fnac"])
    out["fin_eff"] = out["fin"].fillna(today_dt())

    for c in ["cod", "clas_raw", "sexo", "ts", "emp", "area_raw", "cargo", "nac", "lug", "reg"]:
        out[c] = out[c].astype("string").str.strip()
        out.loc[out[c].isin(["", "None", "nan", "NaT"]), c] = pd.NA

    out = out[~out["cod"].isna()].copy()
    out = out[~out["ini"].isna()].copy()
    out["cod"] = out["cod"].astype(str)

    out["area"], out["area_gen"] = _map_area(out["area_raw"])
    out["clas"] = _map_clas(out["clas_raw"])

    out = out[KEEP_INTERNAL].copy()
    out = out.sort_values(["cod", "ini", "fin_eff"]).reset_index(drop=True)
    return out


# =========================
# Merge intervalos por persona (para existencias)
# (ojo: esto une continuidad laboral; NO depende de categorías)
# =========================
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


# =========================
# Filtros
# =========================
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
    antig: List[str]  # bucket (aplica a exposición y salidas en su fecha)
    edad: List[str]   # bucket (aplica a exposición y salidas en su fecha)

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


# =========================
# Buckets seleccionables
# =========================
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


# =========================
# Existencias diarias (optimizado) + buckets (antig/edad)
# =========================
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

        # Antig buckets -> rangos
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

        # Edad buckets -> rangos
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


# =========================
# Salidas diarias (por fin) + buckets en fecha de salida
# =========================
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


# =========================
# Ventanas por periodo (se mantiene para transición / utilidades)
# =========================
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


# =========================
# Agregación diaria -> periodo para PDE
# - Exposicion = suma de existencias (días-persona OBS)
# - Perdidos = sum(Salidas_dia * (T1 - dia))
# - Potencial = Exposicion + Perdidos
# - PDE = Perdidos / Potencial
# - Existencias_Prom = Exposicion / L
# =========================
def aggregate_daily_to_period_for_pde(df_daily_g: pd.DataFrame, period: str) -> pd.DataFrame:
    d = df_daily_g.copy()

    # Asegura calendario
    if "CodSem" not in d.columns or "CodMes" not in d.columns or "Año" not in d.columns:
        d = add_calendar_fields(d, "Día")

    key = {"D": "Día", "W": "CodSem", "M": "CodMes", "Y": "Año"}[period]
    cut_col = {"D": "Día", "W": "FinSemana", "M": "FinMes", "Y": "Día"}[period]

    # Asegura columnas base
    if "Salidas" not in d.columns:
        d["Salidas"] = 0
    if "Existencias" not in d.columns:
        d["Existencias"] = 0

    def _agg_group(g: pd.DataFrame) -> pd.Series:
        ws = g["Día"].min()
        we = g["Día"].max()
        L = int((we - ws).days + 1) if pd.notna(ws) and pd.notna(we) else 0

        expo = float(np.nansum(g["Existencias"].astype(float).values))  # Obs (días-persona)
        sal = float(np.nansum(g["Salidas"].astype(float).values))

        # Perdidos: por cada salida en día d, pierde (T1 - d) días potenciales post-salida
        delta = (we - g["Día"]).dt.days.astype(float)
        perd = float(np.nansum(g["Salidas"].astype(float).values * delta.values))

        pot = expo + perd
        pde = (perd / pot) if pot > 0 else np.nan
        exist_prom = (expo / L) if L > 0 else np.nan

        # metadatos
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

    # Periodo etiqueta
    if period == "D":
        agg["Periodo"] = pd.to_datetime(agg["window_start"]).dt.strftime("%Y-%m-%d")
    elif period in ("W", "M"):
        agg["Periodo"] = agg[key].astype(str) if key in agg.columns else agg["CodSem"].astype(str)
    else:  # "Y"
        agg["Periodo"] = agg[key].astype(int).astype(str) if key in agg.columns else agg["Año"].astype(int).astype(str)

    # Orden
    agg = agg.sort_values("cut").reset_index(drop=True)

    # Sanity check: columnas críticas
    for c in ["Periodo", "cut", "window_start", "window_end", "Exposicion", "Potencial", "PDE", "Existencias_Prom", "Perdidos"]:
        if c not in agg.columns:
            raise ValueError(f"aggregate_daily_to_period_for_pde: falta columna {c}")

    return agg


# =========================
# KPI PDE:
# - PDE_H = 1 - (1 - PDE)^(H / L)
# - Meta(t) = promedio(PDE_H_base(t-S), t-S-1, t-S-2)
# - KPI(t) = PDE_H_g(t) / Meta(t)
# - Mejora(t) = KPI(t-1) - KPI(t)
# - Semáforo: Verde <=0.95, Amarillo <=1.05, Rojo >1.05
# =========================
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
    # tolerancia si Dias_Periodo no viene
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

    # PDE_H (equivalente a horizonte fijo H)
    L = m["Dias_Periodo"].astype(float).replace(0, np.nan)
    m["PDEH_g"] = 1.0 - np.power((1.0 - m["PDE_g"].astype(float)).clip(lower=0.0, upper=1.0), (H / (L + eps)))
    m["PDEH_b"] = 1.0 - np.power((1.0 - m["PDE_b"].astype(float)).clip(lower=0.0, upper=1.0), (H / (L + eps)))

    # Seasonality S
    if period == "W":
        S = 52
    elif period == "M":
        S = 12
    elif period == "D":
        S = 365
    else:  # "Y"
        S = 1
    m["S"] = S

    # Meta estacional suavizada (baseline)
    p = m["PDEH_b"].astype(float)
    m["Meta"] = (p.shift(S) + p.shift(S + 1) + p.shift(S + 2)) / 3.0

    # KPI y mejora
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

    # Regla de acción (operativa)
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

    # Lectura ejecutiva
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


# =========================
# Contingencias
# =========================
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

RECOMMENDED_CROSSES = [
    ("Área General", "Área"),
    ("Área General", "Clasificación"),
    ("Área", "Antigüedad (bucket)"),
    ("Área", "Edad (bucket)"),
    ("Área", "Sexo"),
    ("Área", "Clasificación"),
    ("Cargo Actual", "Sexo"),
    ("Empresa", "Área General"),
    ("Región Registro", "Nacionalidad"),
]

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


# =========================
# Snapshot buckets
# =========================
def add_snapshot_buckets(df_snap: pd.DataFrame, cut: pd.Timestamp) -> pd.DataFrame:
    if df_snap.empty:
        return df_snap
    df = df_snap.copy()
    df["ref"] = cut
    df["antig_dias"] = (df["ref"] - df["ini"]).dt.days
    df["Antigüedad"] = bucket_antiguedad(df["antig_dias"])
    df["Edad"] = bucket_edad_from_dob(df["fnac"], df["ref"])
    return df


# =========================
# Transición (cohorte activa en t1 -> bucket en t2 o SALIO)
# =========================
def transition_table_antig(
    df_intervals_filtered_base: pd.DataFrame,
    cut1: Optional[pd.Timestamp] = None,
    cut2: Optional[pd.Timestamp] = None,
    include_left_as: str = "SALIO",
    top_rows: int = 80,
    top_cols: int = 80,
    fecha1: Optional[pd.Timestamp] = None,
    fecha2: Optional[pd.Timestamp] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if cut1 is None:
        cut1 = fecha1
    if cut2 is None:
        cut2 = fecha2
    if cut1 is None or cut2 is None:
        raise ValueError("transition_table_antig requiere cut1 y cut2 (o fecha1/fecha2).")

    cut1 = pd.Timestamp(cut1).normalize()
    cut2 = pd.Timestamp(cut2).normalize()

    s1 = df_intervals_filtered_base[
        (df_intervals_filtered_base["ini"] <= cut1) & (df_intervals_filtered_base["fin_eff"] >= cut1)
    ].copy()

    if s1.empty:
        return pd.DataFrame(), pd.DataFrame()

    s1 = s1.sort_values(["cod", "fin_eff", "ini"]).drop_duplicates("cod", keep="last")
    s1["ref1"] = cut1
    s1["antig_dias_1"] = (s1["ref1"] - s1["ini"]).dt.days
    s1["Antigüedad_1"] = bucket_antiguedad(s1["antig_dias_1"])
    s1 = s1[["cod", "Antigüedad_1"]].copy()

    s2 = df_intervals_filtered_base[
        (df_intervals_filtered_base["ini"] <= cut2) & (df_intervals_filtered_base["fin_eff"] >= cut2)
    ].copy()

    if not s2.empty:
        s2 = s2.sort_values(["cod", "fin_eff", "ini"]).drop_duplicates("cod", keep="last")
        s2["ref2"] = cut2
        s2["antig_dias_2"] = (s2["ref2"] - s2["ini"]).dt.days
        s2["Antigüedad_2"] = bucket_antiguedad(s2["antig_dias_2"])
        s2 = s2[["cod", "Antigüedad_2"]].copy()
    else:
        s2 = pd.DataFrame({"cod": pd.Series(dtype="object"), "Antigüedad_2": pd.Series(dtype="object")})

    m = s1.merge(s2, on="cod", how="left")
    m["Antigüedad_2"] = m["Antigüedad_2"].fillna(include_left_as)

    r = _norm_cat(m["Antigüedad_1"])
    c = _norm_cat(m["Antigüedad_2"])

    counts = pd.crosstab(r, c, dropna=False)
    counts = _reduce_crosstab(counts, top_rows=top_rows, top_cols=top_cols)

    denom = counts.sum(axis=1).replace(0, np.nan)
    pct = (counts.div(denom, axis=0) * 100).round(2)

    counts.index = counts.index.map(str)
    counts.columns = counts.columns.map(str)
    pct.index = pct.index.map(str)
    pct.columns = pct.columns.map(str)

    return counts, pct


# =========================
# Plot helpers
# =========================
def apply_bar_labels(fig, show_labels: bool, fmt: str = ".0f"):
    if show_labels:
        fig.update_traces(texttemplate=f"%{{y:{fmt}}}", textposition="outside", cliponaxis=False)
    return fig

def apply_line_labels(fig, show_labels: bool, fmt: str = ".2f"):
    if show_labels:
        fig.update_traces(mode="lines+markers+text", texttemplate=f"%{{y:{fmt}}}", textposition="top center")
    return fig


# =========================
# Lectura robusta
# =========================
def read_excel_strict(file_obj_or_path, sheet_name: Optional[str]):
    try:
        return pd.read_excel(file_obj_or_path, sheet_name=sheet_name, usecols=REQUIRED_COLS)
    except Exception:
        return pd.read_excel(file_obj_or_path, sheet_name=sheet_name)

def read_csv_strict(file_obj_or_path):
    try:
        return pd.read_csv(file_obj_or_path, usecols=lambda c: c in REQUIRED_COLS)
    except Exception:
        return pd.read_csv(file_obj_or_path)


# =========================
# Export helpers
# =========================
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLS = 16_384

def df_exceeds_excel_limits(df: pd.DataFrame) -> bool:
    return (df.shape[0] > EXCEL_MAX_ROWS) or (df.shape[1] > EXCEL_MAX_COLS)

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


# =========================
# Layout
# =========================
main_col, filter_col = st.columns([4.3, 1.7], gap="large")


# =========================
# Panel de filtros (derecha)
# =========================
with filter_col:
    st.subheader("Panel de control")

    tab_f1, tab_f2, tab_f3, tab_f4 = st.tabs(["Periodo", "Demografía", "Opciones", "Transición"])

    # ---- Periodo + Carga ----
    with tab_f1:
        st.markdown("**Datos**")
        uploaded = st.file_uploader("Sube Excel/CSV (Historial_Personal)", type=["xlsx", "xls", "csv"], key="uploader_main")
        path = st.text_input("O ruta local (opcional)", value="", key="path_input")

        sheet_name = None
        df_raw = None

        if uploaded is not None:
            try:
                if uploaded.name.lower().endswith(".csv"):
                    df_raw = read_csv_strict(uploaded)
                else:
                    xls = pd.ExcelFile(uploaded)
                    sheet_name = st.selectbox("Hoja", options=xls.sheet_names, index=0, key="sheet_select_upload")
                    df_raw = read_excel_strict(uploaded, sheet_name)
            except Exception as e:
                st.error(f"No se pudo leer el archivo: {e}")
                st.stop()

        elif path.strip():
            if not os.path.exists(path.strip()):
                st.error("La ruta no existe.")
                st.stop()
            try:
                if path.lower().endswith(".csv"):
                    df_raw = read_csv_strict(path.strip())
                else:
                    xls = pd.ExcelFile(path.strip())
                    sheet_name = st.selectbox("Hoja", options=xls.sheet_names, index=0, key="sheet_select_path")
                    df_raw = read_excel_strict(path.strip(), sheet_name)
            except Exception as e:
                st.error(f"No se pudo leer el archivo: {e}")
                st.stop()
        else:
            st.info("Carga un archivo para iniciar.")
            st.stop()

        try:
            df0 = validate_and_prepare(df_raw)
        except Exception as e:
            st.error(str(e))
            st.stop()

        df_intervals_all = merge_intervals_per_person(df0)

        st.markdown("**Rango de análisis**")
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

        use_slider = st.checkbox("Usar deslizador de fechas", value=True, key="use_date_slider")

        if use_slider:
            r0, r1 = st.slider(
                "Inicio / Fin",
                min_value=(min_date.date() if pd.notna(min_date) else date(2000, 1, 1)),
                max_value=(max_date.date() if pd.notna(max_date) else default_end.date()),
                value=st.session_state["date_range_main"],
                key="date_range_slider",
            )
            st.session_state["date_range_main"] = (r0, r1)
        else:
            rango = st.date_input(
                "Inicio / Fin",
                key="date_range_main",
                min_value=(min_date.date() if pd.notna(min_date) else None),
                max_value=(max_date.date() if pd.notna(max_date) else None),
            )
            if not isinstance(rango, tuple) or len(rango) != 2:
                st.error("Selecciona (inicio, fin).")
                st.stop()

        start_dt = pd.Timestamp(st.session_state["date_range_main"][0])
        end_dt = pd.Timestamp(st.session_state["date_range_main"][1])
        if start_dt > end_dt:
            st.error("Inicio > Fin.")
            st.stop()

        period_label = st.selectbox("Agrupar por", options=["Día", "Semana", "Mes", "Año"], index=1, key="period_group")
        period = {"Día": "D", "Semana": "W", "Mes": "M", "Año": "Y"}[period_label]

        snap_date = st.slider(
            "Snapshot de existencias (día)",
            min_value=start_dt.date(),
            max_value=end_dt.date(),
            value=end_dt.date(),
            key="snap_date",
        )
        snap_dt = pd.Timestamp(snap_date)

    # ---- Demografía (dropdown gerencial con TODOS)
    with tab_f2:
        st.markdown("**Filtros demográficos (aplican a todo el dashboard)**")
        st.caption("Modo gerencial: cada filtro es un dropdown con opción 'TODOS'. (Si dejas TODOS, no filtra.)")

        def opts(df: pd.DataFrame, col: str) -> List[str]:
            v = df[col].dropna().astype(str).str.strip()
            v = v[v != ""].unique().tolist()
            return sorted(v)

        def selectbox_todos(label: str, options: List[str], key: str) -> List[str]:
            all_opts = ["TODOS"] + options
            prev = st.session_state.get(key, "TODOS")
            idx = all_opts.index(prev) if prev in all_opts else 0
            sel = st.selectbox(label, options=all_opts, index=idx, key=key)
            return [] if sel == "TODOS" else [sel]

        if st.button("Limpiar filtros demográficos", use_container_width=True, key="btn_clear_filters"):
            for k in [
                "f_sexo", "f_area_gen", "f_area", "f_cargo", "f_clas", "f_ts", "f_emp", "f_nac", "f_lug", "f_reg",
                "f_antig", "f_edad",
            ]:
                st.session_state[k] = "TODOS"
            st.rerun()

        area_gen_pick = selectbox_todos("Área General", opts(df0, "area_gen"), "f_area_gen")
        if area_gen_pick:
            df_area = df0[df0["area_gen"].isin(area_gen_pick)]
            area_opts = opts(df_area, "area")
        else:
            area_opts = opts(df0, "area")

        fs = FilterState(
            sexo=selectbox_todos("Sexo", opts(df0, "sexo"), "f_sexo"),
            area_gen=area_gen_pick,
            area=selectbox_todos("Área (nombre)", area_opts, "f_area"),
            cargo=selectbox_todos("Cargo", opts(df0, "cargo"), "f_cargo"),
            clas=selectbox_todos("Clasificación (nombre)", opts(df0, "clas"), "f_clas"),
            ts=selectbox_todos("Trabajadora Social", opts(df0, "ts"), "f_ts"),
            emp=selectbox_todos("Empresa", opts(df0, "emp"), "f_emp"),
            nac=selectbox_todos("Nacionalidad", opts(df0, "nac"), "f_nac"),
            lug=selectbox_todos("Lugar Registro", opts(df0, "lug"), "f_lug"),
            reg=selectbox_todos("Región Registro", opts(df0, "reg"), "f_reg"),
            antig=selectbox_todos(
                "Antigüedad (bucket) [aplica a exposición y salidas]",
                ["< 30 días", "30 - 90 días", "91 - 180 días", "181 - 360 días", "> 360 días", MISSING_LABEL],
                "f_antig",
            ),
            edad=selectbox_todos(
                "Edad (bucket) [aplica a exposición y salidas]",
                ["< 24 años", "24 - 30 años", "31 - 37 años", "38 - 42 años", "43 - 49 años", "50 - 56 años", "> 56 años", MISSING_LABEL],
                "f_edad",
            ),
        )

    # ---- Opciones ----
    with tab_f3:
        unique_personas_por_dia = st.checkbox("Salidas: contar personas únicas por día", value=True, key="opt_unique_day")
        show_labels = st.checkbox("Mostrar etiquetas numéricas en gráficos", value=True, key="opt_show_labels")

        st.markdown("**Horizonte fijo H (para PDE_H equivalente)**")
        h_choice = st.selectbox("Horizonte", options=["30 días", "90 días", "180 días", "Otro…"], index=0, key="opt_h_choice")
        if h_choice == "30 días":
            horizon_days = 30
        elif h_choice == "90 días":
            horizon_days = 90
        elif h_choice == "180 días":
            horizon_days = 180
        else:
            horizon_days = int(st.number_input("H (días)", min_value=7, max_value=365, value=30, step=1, key="opt_h_custom"))

        st.markdown("**Semáforo (KPI_PDE)**")
        green_max = float(st.number_input("Verde si ≤", min_value=0.70, max_value=1.10, value=0.95, step=0.01, key="opt_green"))
        yellow_max = float(st.number_input("Amarillo si ≤", min_value=0.80, max_value=1.30, value=1.05, step=0.01, key="opt_yellow"))

        st.markdown("**Contingencias**")
        cont_unique_range = st.checkbox("Salidas (contingencia): 1 registro por persona en el rango", value=True, key="opt_cont_unique_range")
        top_rows = st.number_input("Máx. filas (categorías)", min_value=10, max_value=200, value=40, step=5, key="opt_top_rows")
        top_cols = st.number_input("Máx. columnas (categorías)", min_value=10, max_value=200, value=40, step=5, key="opt_top_cols")

        st.caption(
            "PDE mide % de capacidad laboral potencial perdida por deserción en el periodo (basado en días-persona). "
            "KPI_PDE = PDE_H_actual / Meta_estacional_suavizada. (<1 mejor, >1 peor)."
        )

    # ---- Transición
    with tab_f4:
        st.markdown("**Transición (Antigüedad bucket)**")
        st.caption("Elige 2 cortes. Si es Semana/Mes, se toma el FIN de la semana/mes.")

        trans_unit = st.selectbox("Unidad", options=["Día", "Semana", "Mes"], index=1, key="trans_unit")

        cal_range = add_calendar_fields(pd.DataFrame({"Día": pd.date_range(start_dt, end_dt, freq="D")}), "Día")

        if trans_unit == "Día":
            t1 = st.date_input("Fecha 1", value=start_dt.date(), key="t1_day")
            t2 = st.date_input("Fecha 2", value=end_dt.date(), key="t2_day")
            trans_cut1 = pd.Timestamp(t1)
            trans_cut2 = pd.Timestamp(t2)
        elif trans_unit == "Semana":
            w = cal_range.groupby("CodSem", as_index=False).agg(cut=("FinSemana", "max"))
            w = w.sort_values("cut")
            opts_w = w["CodSem"].astype(str).tolist()
            map_w = dict(zip(w["CodSem"].astype(str), w["cut"]))
            pick1 = st.selectbox("Semana 1 (YYWW)", options=opts_w, index=0, key="t1_wk")
            pick2 = st.selectbox("Semana 2 (YYWW)", options=opts_w, index=len(opts_w) - 1, key="t2_wk")
            trans_cut1 = pd.Timestamp(map_w[pick1])
            trans_cut2 = pd.Timestamp(map_w[pick2])
        else:
            m_ = cal_range.groupby("CodMes", as_index=False).agg(cut=("FinMes", "max"))
            m_ = m_.sort_values("cut")
            opts_m = m_["CodMes"].astype(str).tolist()
            map_m = dict(zip(m_["CodMes"].astype(str), m_["cut"]))
            pick1 = st.selectbox("Mes 1 (YYMM)", options=opts_m, index=0, key="t1_mo")
            pick2 = st.selectbox("Mes 2 (YYMM)", options=opts_m, index=len(opts_m) - 1, key="t2_mo")
            trans_cut1 = pd.Timestamp(map_m[pick1])
            trans_cut2 = pd.Timestamp(map_m[pick2])

        if trans_cut1 > trans_cut2:
            st.error("Fecha 1 debe ser <= Fecha 2.")
            st.stop()


# =========================
# Cálculos
# =========================
with st.spinner("Calculando métricas..."):
    # OBS (con filtros)
    df0_f = apply_categorical_filters(df0, fs)

    # Intervalos para exposición (se mergean DESPUÉS de filtrar para evitar distorsión por movimientos)
    df_intervals_f = merge_intervals_per_person(df0_f) if not df0_f.empty else df0_f.copy()

    # ---- Grupo (filtrado)
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

    # ---- Baseline (global, mismo rango, mismos buckets antig/edad si están seleccionados)
    df_salidas_daily_b, _df_sal_det_b = compute_salidas_daily_filtered(
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

    # ---- Periodo (PDE)
    df_period_g = aggregate_daily_to_period_for_pde(df_daily_g, period)
    df_period_b = aggregate_daily_to_period_for_pde(df_daily_b, period)

    # ---- KPI (PDE_H + KPI_PDE vs Meta estacional)
    df_kpi = compute_pde_kpis(
        df_period_g=df_period_g,
        df_period_base=df_period_b,
        horizon_days=int(horizon_days),
        period=period,
        green_max=float(green_max),
        yellow_max=float(yellow_max),
    )

    # Snapshot (existencias a un día)
    df_snap = df_intervals_f[(df_intervals_f["ini"] <= snap_dt) & (df_intervals_f["fin_eff"] >= snap_dt)].copy()
    df_snap = add_snapshot_buckets(df_snap, snap_dt)
    if fs.antig and ("Antigüedad" in df_snap.columns):
        df_snap = df_snap[df_snap["Antigüedad"].isin(fs.antig)]
    if fs.edad and ("Edad" in df_snap.columns):
        df_snap = df_snap[df_snap["Edad"].isin(fs.edad)]

    # Transición
    df_trans_counts, df_trans_pct = transition_table_antig(
        df_intervals_filtered_base=df_intervals_f,
        cut1=trans_cut1,
        cut2=trans_cut2,
        include_left_as="SALIO",
        top_rows=80,
        top_cols=80,
    )


# =========================
# UI Principal
# =========================
with main_col:
    sal_total = int(df_daily_g["Salidas"].sum()) if not df_daily_g.empty else 0
    exist_prom = float(df_daily_g["Existencias"].mean()) if not df_daily_g.empty else 0.0
    exist_snap = int(df_snap["cod"].nunique()) if (df_snap is not None and not df_snap.empty) else 0

    pdeh_last = np.nan
    kpi_last = np.nan
    sem_last = "-"
    accion_last = ""

    if not df_kpi.empty and df_kpi["PDEH_g"].dropna().shape[0] > 0:
        last = df_kpi.dropna(subset=["PDEH_g"]).iloc[-1]
        pdeh_last = float(last["PDEH_g"])
        kpi_last = float(last["KPI_PDE"]) if not np.isnan(float(last.get("KPI_PDE", np.nan))) else np.nan
        sem_last = str(last.get("Semaforo", "-"))
        accion_last = str(last.get("Accion", ""))

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Salidas (rango)", f"{sal_total:,}".replace(",", "."))
    k2.metric("Existencias promedio (rango)", f"{exist_prom:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    k3.metric(f"Existencias snapshot {snap_dt.date()}", f"{exist_snap:,}".replace(",", "."))
    k4.metric(f"PDE{int(horizon_days)} (último, equivalente)", "-" if np.isnan(pdeh_last) else f"{pdeh_last:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
    k5.metric("KPI_PDE (último)", "-" if np.isnan(kpi_last) else f"{kpi_last:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))

    if accion_last == "INTERVENIR":
        st.error("Regla de acción activada: **INTERVENIR** (Rojo 2 seguidos o Amarillo 3 en alza).")

    tabs = st.tabs(["KPI Gerencial (PDE)", "PDE (capacidad perdida)", "Diario", "Contingencias", "Transición", "Descargas"])

    # ---- KPI GERENCIAL
    with tabs[0]:
        st.subheader("KPI gerencial: PDE vs Meta estacional (suavizada)")
        st.caption(
            "PDE mide % de capacidad laboral potencial perdida por deserción, usando días-persona. "
            "PDE_H convierte cualquier periodo a un horizonte fijo H (ej. 30 días) para comparar semana/mes/año."
        )

        if df_kpi.empty or df_kpi["PDEH_g"].dropna().empty:
            st.warning("No hay datos suficientes para calcular PDE/KPI con los filtros actuales.")
        else:
            last = df_kpi.dropna(subset=["PDEH_g"]).iloc[-1]
            pdehg = float(last["PDEH_g"])
            meta = float(last["Meta"]) if pd.notna(last.get("Meta", np.nan)) else np.nan
            kpi = float(last["KPI_PDE"]) if pd.notna(last.get("KPI_PDE", np.nan)) else np.nan
            brecha = float(last["Brecha_vs_Meta"]) if pd.notna(last.get("Brecha_vs_Meta", np.nan)) else np.nan
            sem = str(last.get("Semaforo", "-"))
            lectura = str(last.get("Lectura", "-"))
            accion = str(last.get("Accion", ""))

            c1, c2, c3, c4 = st.columns(4)
            c1.metric(f"PDE{int(horizon_days)} (grupo)", f"{pdehg:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
            c2.metric("Meta estacional (baseline)", "-" if np.isnan(meta) else f"{meta:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
            c3.metric("KPI_PDE = PDE/Meta", "-" if np.isnan(kpi) else f"{kpi:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
            c4.metric("Semáforo", sem)

            st.markdown(
                f"""
**Cómo leerlo (directo para gerencia)**  
- **PDE** = % de capacidad potencial perdida por deserción dentro del periodo.  
- **PDE{int(horizon_days)} (equivalente)** = mismo significado aunque compares semanas vs meses (horizonte fijo H).  
- **Meta estacional** = promedio del mismo periodo del año anterior (suavizado con 3 periodos): t−S, t−S−1, t−S−2.  
- **KPI_PDE = PDE{int(horizon_days)} / Meta**:  
  - KPI = 1.00 → igual a la meta  
  - KPI < 1.00 → **mejor** (menos capacidad perdida)  
  - KPI > 1.00 → **peor** (más capacidad perdida)  
- **Brecha** (PDE−Meta): {("-" if np.isnan(brecha) else f"{brecha:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))}  
- Lectura: **{lectura}**  
- Acción: **{accion if accion else "—"}**
                """
            )

            fig_kpi = px.line(df_kpi, x="Periodo", y="KPI_PDE", title="KPI_PDE (PDE_H / Meta)")
            fig_kpi.update_xaxes(type="category")
            st.plotly_chart(apply_line_labels(fig_kpi, show_labels, ".3f"), use_container_width=True)

            fig_pde = px.line(df_kpi, x="Periodo", y="PDEH_g", title=f"PDE{int(horizon_days)} (grupo)")
            fig_pde.update_xaxes(type="category")
            st.plotly_chart(apply_line_labels(fig_pde, show_labels, ".3f"), use_container_width=True)

            fig_meta = px.line(df_kpi, x="Periodo", y="Meta", title="Meta estacional suavizada (baseline)")
            fig_meta.update_xaxes(type="category")
            st.plotly_chart(apply_line_labels(fig_meta, show_labels, ".3f"), use_container_width=True)

            fig_mej = px.bar(df_kpi, x="Periodo", y="Mejora", title="Mejora(t) = KPI(t−1) − KPI(t)")
            fig_mej.update_xaxes(type="category")
            st.plotly_chart(apply_bar_labels(fig_mej, show_labels, ".3f"), use_container_width=True)

            st.dataframe(
                _safe_table_for_streamlit(
                    df_kpi[
                        ["Periodo", "Dias_Periodo",
                         "Obs_g", "Perdidos_g", "Pot_g", "PDE_g", "PDEH_g",
                         "Meta", "KPI_PDE", "Mejora", "Semaforo", "Accion"]
                    ].tail(60)
                ),
                use_container_width=True,
                height=430,
            )

    # ---- PDE (capacidad perdida)
    with tabs[1]:
        st.subheader(f"PDE (capacidad laboral perdida) por {period_label}")

        st.markdown(
            f"""
### Interpretación (simple y robusta)
- **Exposición (Obs)** = días-persona observados (suma de existencias diarias).  
- **Potencial** = Obs + Perdidos.  
- **Perdidos** = días-persona que se pierden por salidas antes del fin del periodo (post-salida).  
- **PDE = Perdidos / Potencial** ⇒ % de capacidad potencial perdida por deserción.  
- **PDE{int(horizon_days)}** = equivalente a horizonte fijo H:  1 − (1−PDE)^(H/L).  

Esto no se distorsiona por:
- tamaño de dotación (500 vs 750),
- granularidad (semana/mes),
- cambios dentro del periodo (porque suma días-persona).
            """
        )

        if df_kpi.empty or df_kpi["PDEH_g"].dropna().empty:
            st.warning("Sin datos para PDE con los filtros actuales.")
        else:
            fig = px.line(df_kpi, x="Periodo", y="PDEH_g", title=f"PDE{int(horizon_days)} (equivalente) en el tiempo")
            fig.update_xaxes(type="category")
            st.plotly_chart(apply_line_labels(fig, show_labels, ".3f"), use_container_width=True)

            fig2 = px.line(df_kpi, x="Periodo", y="PDE_g", title="PDE (crudo del periodo)")
            fig2.update_xaxes(type="category")
            st.plotly_chart(apply_line_labels(fig2, show_labels, ".3f"), use_container_width=True)

            st.markdown("---")
            st.subheader("Complemento operativo: Perdidos, Potencial, Exposición y Existencias promedio")
            fig_per = px.bar(df_period_g, x="Periodo", y="Perdidos", title="Días-persona perdidos por periodo")
            fig_per.update_xaxes(type="category")
            st.plotly_chart(apply_bar_labels(fig_per, show_labels, ".0f"), use_container_width=True)

            fig_pot = px.bar(df_period_g, x="Periodo", y="Potencial", title="Días-persona potenciales por periodo")
            fig_pot.update_xaxes(type="category")
            st.plotly_chart(apply_bar_labels(fig_pot, show_labels, ".0f"), use_container_width=True)

            fig_obs = px.bar(df_period_g, x="Periodo", y="Exposicion", title="Exposición observada por periodo (días-persona)")
            fig_obs.update_xaxes(type="category")
            st.plotly_chart(apply_bar_labels(fig_obs, show_labels, ".0f"), use_container_width=True)

            fig_ex_p = px.line(df_period_g, x="Periodo", y="Existencias_Prom", title="Existencias promedio por periodo (personas)")
            fig_ex_p.update_xaxes(type="category")
            st.plotly_chart(apply_line_labels(fig_ex_p, show_labels, ".1f"), use_container_width=True)

    # ---- Diario
    with tabs[2]:
        st.subheader("Tabla diaria (grupo filtrado)")
        view = df_daily_g[["Día", "Salidas", "Existencias", "Semana", "Mes", "Año", "CodSem", "CodMes"]].copy()
        view["CodSem"] = view["CodSem"].astype(str)
        view["CodMes"] = view["CodMes"].astype(str)

        st.dataframe(_safe_table_for_streamlit(view), use_container_width=True, height=520)

        fig_e = px.line(view, x="Día", y="Existencias", title="Existencias diarias")
        fig_e.update_layout(yaxis_title="Existencias (personas)")
        st.plotly_chart(apply_line_labels(fig_e, show_labels, ".0f"), use_container_width=True)

        fig_s = px.bar(view, x="Día", y="Salidas", title="Salidas diarias")
        fig_s.update_layout(yaxis_title="Salidas (conteo)")
        st.plotly_chart(apply_bar_labels(fig_s, show_labels, ".0f"), use_container_width=True)

    # ---- Contingencias
    with tabs[3]:
        st.subheader("Tablas de contingencia (% por fila)")
        sub_tabs = st.tabs(["Salidas (detalle en rango)", f"Existencias (snapshot {snap_dt.date()})"])

        def contingency_ui(df_ctx: pd.DataFrame, default_pair: Tuple[str, str], key_base: str, is_salidas: bool) -> None:
            if df_ctx is None or df_ctx.empty:
                st.warning("No hay registros para contingencias con los filtros actuales.")
                return

            df_use = df_ctx.copy()

            if is_salidas and cont_unique_range and ("cod" in df_use.columns):
                if "ref_fin" in df_use.columns:
                    df_use = df_use.sort_values("ref_fin").drop_duplicates("cod", keep="last")
                else:
                    df_use = df_use.drop_duplicates("cod", keep="last")

            pairs = [f"{a} vs {b}" for a, b in RECOMMENDED_CROSSES]
            default_label = f"{default_pair[0]} vs {default_pair[1]}"
            default_idx = pairs.index(default_label) if default_label in pairs else 0

            pick = st.selectbox("Cruce recomendado", options=pairs, index=default_idx, key=f"{key_base}_pick")
            row_dim, col_dim = pick.split(" vs ")

            with st.expander("Cambiar cruce manualmente", expanded=False):
                row_dim = st.selectbox("Filas", options=list(DIMENSIONS.keys()), index=list(DIMENSIONS.keys()).index(row_dim), key=f"{key_base}_rows")
                col_dim = st.selectbox("Columnas", options=list(DIMENSIONS.keys()), index=list(DIMENSIONS.keys()).index(col_dim), key=f"{key_base}_cols")

            counts, pct = contingency_tables(df_use, row_dim, col_dim, top_rows=int(top_rows), top_cols=int(top_cols))

            counts_show = counts.reset_index().rename(columns={"index": row_dim})
            pct_show = pct.reset_index().rename(columns={"index": row_dim})

            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Conteos**")
                st.dataframe(_safe_table_for_streamlit(counts_show), use_container_width=True, height=420)
            with c2:
                st.markdown("**% por fila**")
                st.dataframe(_safe_table_for_streamlit(pct_show), use_container_width=True, height=420)

            fig = px.imshow(
                pct.values,
                x=[str(x) for x in pct.columns],
                y=[str(y) for y in pct.index],
                aspect="auto",
                text_auto=True,
                title=f"Heatmap (% fila): {row_dim} vs {col_dim}",
            )
            fig.update_layout(coloraxis_colorbar_title="% fila")
            st.plotly_chart(fig, use_container_width=True)

            st.caption("Cada fila suma 100%.")

        with sub_tabs[0]:
            contingency_ui(df_sal_det, default_pair=("Área", "Antigüedad (bucket)"), key_base="salidas", is_salidas=True)

        with sub_tabs[1]:
            contingency_ui(df_snap, default_pair=("Área General", "Clasificación"), key_base="snap", is_salidas=False)

    # ---- Transición
    with tabs[4]:
        st.subheader("Transición de Antigüedad (bucket)")
        st.caption(f"Cohorte base: activos en {trans_cut1.date()}. Luego: bucket en {trans_cut2.date()} (o SALIO).")

        if df_trans_counts.empty:
            st.warning("No hay cohorte activa en la Fecha 1 con los filtros actuales.")
        else:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Conteos**")
                st.dataframe(
                    _safe_table_for_streamlit(df_trans_counts.reset_index().rename(columns={"index": "Antigüedad_t1"})),
                    use_container_width=True,
                    height=420,
                )
            with c2:
                st.markdown("**% por fila**")
                st.dataframe(
                    _safe_table_for_streamlit(df_trans_pct.reset_index().rename(columns={"index": "Antigüedad_t1"})),
                    use_container_width=True,
                    height=420,
                )

            figt = px.imshow(
                df_trans_pct.values,
                x=[str(x) for x in df_trans_pct.columns],
                y=[str(y) for y in df_trans_pct.index],
                aspect="auto",
                text_auto=True,
                title="Heatmap transición (% fila)",
            )
            figt.update_layout(coloraxis_colorbar_title="% fila")
            st.plotly_chart(figt, use_container_width=True)

    # ---- Descargas
    with tabs[5]:
        st.subheader("Exportar resultados")

        buf_xlsx = io.BytesIO()
        excel_ok = True
        try:
            with pd.ExcelWriter(buf_xlsx, engine="openpyxl") as writer:
                df_kpi.to_excel(writer, index=False, sheet_name="KPI_PDE")
                df_period_g.to_excel(writer, index=False, sheet_name="Periodo_Grupo")
                df_period_b.to_excel(writer, index=False, sheet_name="Periodo_Baseline")
                df_daily_g[["Día", "Salidas", "Existencias", "Semana", "Mes", "Año", "CodSem", "CodMes"]].to_excel(writer, index=False, sheet_name="Diario_Grupo")
                df_snap.to_excel(writer, index=False, sheet_name="Existencias_Snapshot")

                if not df_trans_counts.empty:
                    df_trans_counts.to_excel(writer, index=True, sheet_name="Transicion_Conteos")
                    df_trans_pct.to_excel(writer, index=True, sheet_name="Transicion_PctFila")

                if df_exceeds_excel_limits(df_sal_det):
                    excel_ok = False
                else:
                    df_sal_det.to_excel(writer, index=False, sheet_name="Salidas_Detalle")
        except Exception:
            excel_ok = False

        if excel_ok:
            st.download_button(
                "Descargar Excel (tablas)",
                data=buf_xlsx.getvalue(),
                file_name="dashboard_pde_resultados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_excel",
            )
        else:
            st.warning("El Excel no pudo incluir 'Salidas_Detalle' (o excede límites). Se exporta Salidas_Detalle como CSV.")
            buf_min = io.BytesIO()
            with pd.ExcelWriter(buf_min, engine="openpyxl") as writer:
                df_kpi.to_excel(writer, index=False, sheet_name="KPI_PDE")
                df_period_g.to_excel(writer, index=False, sheet_name="Periodo_Grupo")
                df_period_b.to_excel(writer, index=False, sheet_name="Periodo_Baseline")
                df_daily_g[["Día", "Salidas", "Existencias", "Semana", "Mes", "Año", "CodSem", "CodMes"]].to_excel(writer, index=False, sheet_name="Diario_Grupo")
                df_snap.to_excel(writer, index=False, sheet_name="Existencias_Snapshot")
                if not df_trans_counts.empty:
                    df_trans_counts.to_excel(writer, index=True, sheet_name="Transicion_Conteos")
                    df_trans_pct.to_excel(writer, index=True, sheet_name="Transicion_PctFila")

            st.download_button(
                "Descargar Excel (sin Salidas_Detalle)",
                data=buf_min.getvalue(),
                file_name="dashboard_pde_resumen.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_excel_min",
            )

            st.download_button(
                "Descargar Salidas_Detalle (CSV)",
                data=to_csv_bytes(df_sal_det),
                file_name="salidas_detalle.csv",
                mime="text/csv",
                key="download_salidas_csv",
            )

st.divider()
st.write(
    "OK: PDE (capacidad laboral perdida por deserción, basado en días-persona) + PDE_H (horizonte fijo) + KPI_PDE vs meta estacional suavizada, "
    "filtros estilo gerencial (TODOS), buckets aplican a exposición y a salidas en su fecha, "
    "y todo lo demás (diario/contingencias/transición/descargas) se mantiene estable."
)
