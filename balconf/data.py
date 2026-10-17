"""Read SAP G/L line-item dumps and build one net balance per vendor."""
from __future__ import annotations

import io
import re

import numpy as np
import pandas as pd

COL_VENDOR = "Vendor"
COL_NAME = "Vendor Account: Name 1"
COL_AMOUNT = "Company Code Currency Value"
REQUIRED = [COL_VENDOR, COL_NAME, COL_AMOUNT]


def _clean_code(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .replace({"nan": "", "None": ""})
    )


def read_dump(content: bytes) -> pd.DataFrame:
    """Read one SAP dump (header in row 1, grand-total row has no vendor)."""
    df = pd.read_excel(io.BytesIO(content), dtype={COL_VENDOR: str})
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Column(s) not found in file: {', '.join(missing)}")
    df[COL_VENDOR] = _clean_code(df[COL_VENDOR])
    df = df[df[COL_VENDOR] != ""].copy()
    df[COL_AMOUNT] = pd.to_numeric(df[COL_AMOUNT], errors="coerce").fillna(0.0)
    return df


def build_balances(adv: pd.DataFrame, cred: pd.DataFrame) -> pd.DataFrame:
    """Net balance per vendor. Positive = debit (advance due from vendor),
    negative = credit (payable to vendor)."""
    a = adv.groupby(COL_VENDOR).agg(name_a=(COL_NAME, "first"), advance=(COL_AMOUNT, "sum"))
    c = cred.groupby(COL_VENDOR).agg(name_c=(COL_NAME, "first"), creditor=(COL_AMOUNT, "sum"))
    df = a.join(c, how="outer")
    df["vendor_name"] = df["name_c"].fillna(df["name_a"]).fillna("").astype(str).str.strip()
    df["advance"] = df["advance"].fillna(0.0).round(2)
    df["creditor"] = df["creditor"].fillna(0.0).round(2)
    df["net"] = (df["advance"] + df["creditor"]).round(2)
    df["nature"] = np.select(
        [df["net"] > 0, df["net"] < 0], ["Debit (due from vendor)", "Credit (payable to vendor)"], "Nil"
    )
    df = df.reset_index().rename(columns={COL_VENDOR: "vendor_code"})
    return df[["vendor_code", "vendor_name", "advance", "creditor", "net", "nature"]]


def _find_col(cols: list[str], *keywords: str, exclude: tuple[str, ...] = ()) -> str | None:
    for c in cols:
        low = c.lower()
        if any(k in low for k in keywords) and not any(x in low for x in exclude):
            return c
    return None


def _norm_emails(v) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    parts = [p.strip() for p in re.split(r"[;,\s]+", str(v)) if p.strip()]
    return "; ".join(parts)


def read_vendor_master(content: bytes) -> pd.DataFrame:
    """Vendor master with columns like 'Vendor Code', 'Email', 'CC' (CC optional)."""
    df = pd.read_excel(io.BytesIO(content), dtype=str)
    cols = [str(c).strip() for c in df.columns]
    df.columns = cols
    code_col = _find_col(cols, "vendor code", "vendor", "code", "supplier")
    email_col = _find_col(cols, "email", "e-mail", "mail", exclude=("cc",))
    cc_col = _find_col(cols, "cc")
    if not code_col or not email_col:
        raise ValueError("Vendor master needs a 'Vendor Code' and an 'Email' column.")
    out = pd.DataFrame({
        "vendor_code": _clean_code(df[code_col]),
        "email": df[email_col].map(_norm_emails),
        "cc": df[cc_col].map(_norm_emails) if cc_col else "",
    })
    out = out[out["vendor_code"] != ""]
    return out.drop_duplicates("vendor_code", keep="last")


def is_valid_emails(s: str) -> bool:
    if not s:
        return False
    return all(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", p.strip()) for p in s.split(";") if p.strip())
