# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import timedelta
from typing import Dict, Tuple
import numpy as np
import pandas as pd

from rrhh_panel.schema.historia_personal import MISSING_LABEL
from rrhh_panel.utils.dates import years_offset_date

TENURE_BUCKETS: Dict[str, Tuple[int, int | None]] = {
    "1) < 30 días": (0, 29),
    "2) 30 - 90 días": (30, 90),
    "3) 91 - 180 días": (91, 180),
    "4) 181 - 360 días": (181, 360),
    "5) > 360 días": (361, None),
}

AGE_BUCKETS: Dict[str, Tuple[int | None, int | None]] = {
    "1) < 24 años": (None, 23),
    "2) 24 - 30 años": (24, 30),
    "3) 31 - 37 años": (31, 37),
    "4) 38 - 42 años": (38, 42),
    "5) 43 - 49 años": (43, 49),
    "6) 50 - 56 años": (50, 56),
    "7) > 56 años": (57, None),
}

def bucket_antiguedad(days: pd.Series) -> pd.Series:
    d = days.astype("float")
    out = pd.Series(np.where(d.isna(), MISSING_LABEL, ""), index=days.index, dtype="object")
    out = np.where((~pd.isna(d)) & (d >= 0) & (d < 30), "1) < 30 días", out)
    out = np.where((~pd.isna(d)) & (d >= 30) & (d <= 90), "2) 30 - 90 días", out)
    out = np.where((~pd.isna(d)) & (d >= 91) & (d <= 180), "3) 91 - 180 días", out)
    out = np.where((~pd.isna(d)) & (d >= 181) & (d <= 360), "4) 181 - 360 días", out)
    out = np.where((~pd.isna(d)) & (d >= 361), "5) > 360 días", out)
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

    out.loc[mask] = np.where(edad < 24, "1) < 24 años", out.loc[mask])
    out.loc[mask] = np.where((edad >= 24) & (edad <= 30), "2) 24 - 30 años", out.loc[mask])
    out.loc[mask] = np.where((edad >= 31) & (edad <= 37), "3) 31 - 37 años", out.loc[mask])
    out.loc[mask] = np.where((edad >= 38) & (edad <= 42), "4) 38 - 42 años", out.loc[mask])
    out.loc[mask] = np.where((edad >= 43) & (edad <= 49), "5) 43 - 49 años", out.loc[mask])
    out.loc[mask] = np.where((edad >= 50) & (edad <= 56), "6) 50 - 56 años", out.loc[mask])
    out.loc[mask] = np.where((edad >= 57), "7) > 56 años", out.loc[mask])
    return out

def make_stratum(active: pd.DataFrame, cut: pd.Timestamp) -> pd.DataFrame:
    a = active.copy()
    a["ref"] = pd.Timestamp(cut).normalize()
    a["tenure_days"] = (a["ref"] - a["ini"]).dt.days
    a["Antigüedad"] = bucket_antiguedad(a["tenure_days"])
    a["Edad"] = bucket_edad_from_dob(a["fnac"], a["ref"])
    a["Estrato"] = a["Edad"].astype(str) + " | " + a["Antigüedad"].astype(str)
    return a
