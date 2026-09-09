from __future__ import annotations

import re
import pandas as pd


def safe_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series,
        errors="coerce"
    ).fillna(0.0)


def normalize_mawb(value) -> str:
    if value is None or pd.isna(value):
        return ""

    text = re.sub(
        r"[^0-9A-Z]",
        "",
        str(value).strip().upper()
    )

    if not text or text in {"NAN", "NONE"}:
        return ""

    if len(text) == 11:
        return f"{text[:3]}-{text[3:]}"

    return str(value).strip().upper()


def parse_mawb_list(text: str) -> list[str]:
    if not text or not str(text).strip():
        return []

    parts = re.split(
        r"[,\s]+",
        str(text).strip()
    )

    return sorted({
        normalize_mawb(x)
        for x in parts
        if normalize_mawb(x)
    })


def pct(
    numerator: pd.Series,
    denominator: pd.Series
) -> pd.Series:

    numerator = pd.to_numeric(
        numerator,
        errors="coerce"
    ).fillna(0.0)

    denominator = pd.to_numeric(
        denominator,
        errors="coerce"
    ).fillna(0.0)

    return (
        numerator / denominator
    ).where(
        denominator.ne(0),
        0.0
    )


def clean_eta_series(
    series: pd.Series
) -> pd.Series:

    return pd.to_datetime(
        series,
        errors="coerce"
    ).dt.normalize()


def display_df(
    df: pd.DataFrame,
    date_cols: list[str] | None = None
) -> pd.DataFrame:

    out = df.copy()

    for col in date_cols or []:
        if col in out.columns:
            out[col] = pd.to_datetime(
                out[col],
                errors="coerce"
            ).dt.date

    for col in [
        c for c in out.columns
        if "%" in str(c)
    ]:

        out[col] = out[col].apply(
            lambda x:
            f"{float(x) * 100:.2f}%"
            if pd.notna(x)
            else ""
        )

    return out
