from __future__ import annotations
import re
import pandas as pd


def safe_numeric(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors='coerce').fillna(0.0)


def norm_colname(s: str) -> str:
    return re.sub(r'[\s_\-]+', '', str(s).strip().lower())


def find_first_col(df: pd.DataFrame, candidates: list[str]) -> str:
    mapping = {norm_colname(c): c for c in df.columns.astype(str)}
    for cand in candidates:
        if norm_colname(cand) in mapping:
            return mapping[norm_colname(cand)]
    return ''


def find_sheet_with_required_cols(xls: pd.ExcelFile, required: dict[str, list[str]]) -> str:
    for sh in xls.sheet_names:
        try:
            tmp = pd.read_excel(xls, sheet_name=sh, nrows=60)
        except Exception:
            continue
        if all(find_first_col(tmp, v) for v in required.values()):
            return sh
    return ''


def pct(numer: pd.Series, denom: pd.Series) -> pd.Series:
    numer = pd.to_numeric(numer, errors='coerce').fillna(0.0)
    denom = pd.to_numeric(denom, errors='coerce').fillna(0.0)
    return (numer / denom).where(denom.ne(0), 0.0)


def normalize_mawb(x) -> str:
    if x is None or pd.isna(x):
        return ''
    s = str(x).strip().upper()
    if not s or s in {'NAN', 'NONE'}:
        return ''
    s = re.sub(r'[^0-9A-Z]', '', s)
    if s.isdigit() and len(s) == 12:
        s = s[-11:]
    if len(s) == 11:
        return f'{s[:3]}-{s[3:]}'
    return s


def parse_mawb_list(text: str) -> list[str]:
    if not text or not str(text).strip():
        return []
    vals = [normalize_mawb(x) for x in re.split(r'[,\s]+', str(text).strip())]
    return sorted({x for x in vals if x})


def clean_eta_series(s: pd.Series) -> pd.Series:
    s = s.astype(str).fillna('').str.strip()
    s = s.str.replace(r'(?i)^\s*eta\s*[:\-]\s*', '', regex=True)
    s = s.str.replace(r'\s+', ' ', regex=True)
    mask = s.str.match(r'^\d{8}$')
    s2 = s.copy()
    if mask.any():
        s2.loc[mask] = pd.to_datetime(s.loc[mask], format='%Y%m%d', errors='coerce').astype(str)
    dt = pd.to_datetime(s2, errors='coerce')
    missing = dt.isna() & s2.ne('')
    if missing.any():
        dt.loc[missing] = pd.to_datetime(s2.loc[missing], errors='coerce', dayfirst=True)
    return dt.dt.normalize()


def to_date_only(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors='coerce').dt.date
    return out


def format_pct_str(x) -> str:
    try:
        return f'{float(x) * 100:.2f}%'
    except Exception:
        return ''


def display_df(df: pd.DataFrame, date_cols: list[str] | None = None) -> pd.DataFrame:
    out = df.copy()
    if date_cols:
        out = to_date_only(out, date_cols)
    for c in ['Profit Margin %', 'Exception %', 'Exempt %', 'Normal %', 'Overall Profit Margin %']:
        if c in out.columns:
            out[c] = out[c].apply(format_pct_str)
    return out
