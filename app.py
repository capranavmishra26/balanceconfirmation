"""Vendor Balance Confirmation - Streamlit app.

Run with:  run.bat   (or:  .venv\\Scripts\\python -m streamlit run app.py)
"""
from __future__ import annotations

import datetime as dt
import io
import json
import re
import time
from pathlib import Path

import pandas as pd
import streamlit as st

from balconf.data import build_balances, is_valid_emails, read_dump, read_vendor_master
from balconf.fmt import inr, long_date, long_date_html
from balconf.letter import build_letter_pdf

ROOT = Path(__file__).parent
CONFIG_PATH = ROOT / "config.json"
OUTPUT_DIR = ROOT / "output"
LOG_PATH = OUTPUT_DIR / "sent_log.csv"

st.set_page_config(page_title="Vendor Balance Confirmation", page_icon="📨", layout="wide")


class _SafeDict(dict):
    def __missing__(self, key):
        return "{" + key + "}"


def fill(template: str, values: dict) -> str:
    return template.format_map(_SafeDict(values))


def load_config() -> dict:
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


@st.cache_data(show_spinner=False)
def cached_dump(content: bytes):
    return read_dump(content)


@st.cache_data(show_spinner=False)
def cached_master(content: bytes):
    return read_vendor_master(content)


def safe_filename(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._ -]+", "", s).strip()[:60] or "vendor"


def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def nature_text(net: float, cfg: dict) -> str:
    if net < 0:
        return cfg["email_nature_credit"]
    if net > 0:
        return cfg["email_nature_debit"]
    return cfg["email_nature_nil"]


def mail_values(row, as_at: dt.date, cfg: dict) -> dict:
    return {
        "vendor_code": row["vendor_code"],
        "vendor_name": row["vendor_name"],
        "as_at": long_date(as_at),
        "as_at_html": long_date_html(as_at),
        "amount": inr(row["net"]),
        "nature": nature_text(row["net"], cfg),
        "company_name": cfg["company_name"],
    }


def append_log(rows: list[dict]) -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    df = pd.DataFrame(rows)
    df.to_csv(LOG_PATH, mode="a", header=not LOG_PATH.exists(), index=False, encoding="utf-8-sig")


cfg = load_config()

# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.header("Settings")
    as_at = st.date_input("Balance as at", value=dt.date(2026, 9, 30), format="DD-MM-YYYY")
    letter_date = st.date_input("Letter date", value=dt.date.today(), format="DD-MM-YYYY")
    min_bal = st.number_input("Minimum balance (absolute, ₹)", min_value=0.0, value=1.0, step=100.0,
                              help="Vendors with a net balance below this are not selected.")
    st.divider()
    st.subheader("Outlook")
    if "accounts" not in st.session_state:
        from balconf.outlook import list_accounts
        st.session_state.accounts = list_accounts()
    accounts = st.session_state.accounts
    if accounts:
        from_account = st.selectbox("Send from", accounts)
    else:
        st.warning("Outlook accounts could not be read. Mail will go from the default account.")
        from_account = None
    test_to = st.text_input("Test mode: send ALL mails to this address",
                            help="Leave blank for live sending to vendors. Vendor CCs are dropped in test mode.")
    delay = st.number_input("Pause between mails (seconds)", min_value=0.0, value=1.0, step=0.5)

st.title("📨 Vendor Balance Confirmation")
st.caption("Upload SAP dumps → review net balances → send PDF confirmation letters through Outlook.")

# ---------------------------------------------------------------- step 1: upload
st.subheader("1. Upload files")
c1, c2, c3 = st.columns(3)
adv_file = c1.file_uploader("Advance to Vendor open items (SAP)", type=["xlsx", "xls"])
cred_file = c2.file_uploader("Creditor open items (SAP)", type=["xlsx", "xls"])
master_file = c3.file_uploader("Vendor master (Vendor Code, Email, CC)", type=["xlsx", "xls"])

if not (adv_file and cred_file):
    st.info("Upload both SAP dumps to continue. The vendor master can be added after downloading the "
            "template in step 2.")
    st.stop()

try:
    with st.spinner("Reading SAP dumps..."):
        adv = cached_dump(adv_file.getvalue())
        cred = cached_dump(cred_file.getvalue())
except ValueError as exc:
    st.error(str(exc))
    st.stop()

bal = build_balances(adv, cred)

if master_file:
    try:
        master = cached_master(master_file.getvalue())
    except ValueError as exc:
        st.error(str(exc))
        st.stop()
    bal = bal.merge(master, on="vendor_code", how="left")
else:
    bal["email"], bal["cc"] = "", ""
bal[["email", "cc"]] = bal[["email", "cc"]].fillna("")
bal["email_ok"] = bal["email"].map(is_valid_emails)

# ---------------------------------------------------------------- step 2: review
st.subheader("2. Review balances")
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Advance lines / vendors", f"{len(adv):,} / {adv['Vendor'].nunique():,}")
m2.metric("Creditor lines / vendors", f"{len(cred):,} / {cred['Vendor'].nunique():,}")
m3.metric("Total vendors", f"{len(bal):,}")
m4.metric("With valid email", f"{int(bal['email_ok'].sum()):,}")
m5.metric("Net balance (₹)", inr(bal["net"].sum()) + (" Cr" if bal["net"].sum() < 0 else " Dr"))

d1, d2 = st.columns(2)
template = bal[["vendor_code", "vendor_name", "email", "cc"]].rename(columns={
    "vendor_code": "Vendor Code", "vendor_name": "Vendor Name", "email": "Email", "cc": "CC"})
d1.download_button("⬇ Vendor master template (fill emails)", to_excel(template),
                   "vendor_master.xlsx", use_container_width=True)
d2.download_button("⬇ Balances (Excel)", to_excel(bal.drop(columns="email_ok")),
                   f"balances_{as_at:%Y%m%d}.xlsx", use_container_width=True)

flt = st.radio("Show", ["All", "Selected-eligible", "Missing email", "Debit (advance)", "Credit"],
               horizontal=True)
bal["send"] = bal["email_ok"] & (bal["net"].abs() >= max(min_bal, 0.01))

view = bal
if flt == "Selected-eligible":
    view = bal[bal["send"]]
elif flt == "Missing email":
    view = bal[~bal["email_ok"]]
elif flt == "Debit (advance)":
    view = bal[bal["net"] > 0]
elif flt == "Credit":
    view = bal[bal["net"] < 0]

edited = st.data_editor(
    view[["send", "vendor_code", "vendor_name", "advance", "creditor", "net", "nature", "email", "cc"]],
    key=f"editor_{flt}_{min_bal}",
    hide_index=True,
    use_container_width=True,
    height=420,
    disabled=["vendor_code", "vendor_name", "advance", "creditor", "net", "nature"],
    column_config={
        "send": st.column_config.CheckboxColumn("Send", width="small"),
        "vendor_code": "Code",
        "vendor_name": st.column_config.TextColumn("Vendor", width="large"),
        "advance": st.column_config.NumberColumn("Advance (Dr)", format="%.2f"),
        "creditor": st.column_config.NumberColumn("Creditor (Cr)", format="%.2f"),
        "net": st.column_config.NumberColumn("Net", format="%.2f"),
        "email": st.column_config.TextColumn("Email", width="medium"),
        "cc": st.column_config.TextColumn("CC", width="medium"),
    },
)
st.caption("Tick/untick **Send**, or type/correct an email directly in the table. "
           "Positive net = debit (advance due from vendor), negative = credit (payable to vendor).")

# Selection = rows ticked in the current view that have a valid email.
selected = edited[edited["send"]].copy()
selected["email_ok"] = selected["email"].map(is_valid_emails)
bad = selected[~selected["email_ok"]]
selected = selected[selected["email_ok"]]
if len(bad):
    st.warning(f"{len(bad)} ticked vendor(s) have no valid email and will be skipped.")

# ---------------------------------------------------------------- step 3: preview
st.subheader("3. Preview")
subject_tpl = st.text_input("Email subject", cfg["email_subject"])
body_tpl = st.text_area("Email body (HTML)", cfg["email_body"], height=150,
                        help="Placeholders: {vendor_name} {vendor_code} {as_at} {amount} {nature} {company_name}")

pool = selected if len(selected) else edited
if len(pool):
    labels = (pool["vendor_code"] + " - " + pool["vendor_name"]).tolist()
    pick = st.selectbox("Preview vendor", range(len(labels)), format_func=lambda i: labels[i])
    row = pool.iloc[pick]
    vals = mail_values(row, as_at, cfg)
    p1, p2 = st.columns([3, 2])
    with p1:
        st.markdown(f"**To:** {test_to or row['email'] or '—'}  \n"
                    f"**CC:** {'' if test_to else row['cc']}  \n"
                    f"**Subject:** {fill(subject_tpl, vals)}")
        st.html(fill(body_tpl, vals))
    with p2:
        prev_path = OUTPUT_DIR / "_preview" / f"BC_{row['vendor_code']}.pdf"
        build_letter_pdf(prev_path, vendor_code=row["vendor_code"], vendor_name=row["vendor_name"],
                         net=row["net"], as_at=as_at, letter_date=letter_date, cfg=cfg)
        st.download_button("⬇ Download preview letter (PDF)", prev_path.read_bytes(),
                           prev_path.name, mime="application/pdf", use_container_width=True)

# ---------------------------------------------------------------- step 4: send
st.subheader("4. Send")
mode = f"TEST mode → all mails go to {test_to}" if test_to else "LIVE mode → mails go to vendors"
st.markdown(f"**{len(selected):,}** vendor(s) selected · {mode}")
confirm = st.checkbox(f"I have reviewed the list and want to send {len(selected):,} email(s) now.")

if st.button("📤 Send emails", type="primary", disabled=not (confirm and len(selected))):
    from balconf.outlook import Outlook

    run_dir = OUTPUT_DIR / f"{as_at:%Y-%m-%d}" / dt.datetime.now().strftime("run_%Y%m%d_%H%M%S")
    progress = st.progress(0.0, text="Starting...")
    results = []
    with Outlook() as ol:
        for i, (_, row) in enumerate(selected.iterrows(), start=1):
            vals = mail_values(row, as_at, cfg)
            pdf = run_dir / f"BC_{row['vendor_code']}_{safe_filename(row['vendor_name'])}.pdf"
            to, cc = (test_to, "") if test_to else (row["email"], row["cc"])
            status, error = "Sent", ""
            try:
                build_letter_pdf(pdf, vendor_code=row["vendor_code"], vendor_name=row["vendor_name"],
                                 net=row["net"], as_at=as_at, letter_date=letter_date, cfg=cfg)
                ol.send(to=to, cc=cc, subject=fill(subject_tpl, vals), html_body=fill(body_tpl, vals),
                        attachments=[pdf], from_account=from_account)
            except Exception as exc:  # keep going; failures are logged
                status, error = "Failed", str(exc)
            results.append({
                "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                "as_at": as_at.isoformat(), "vendor_code": row["vendor_code"],
                "vendor_name": row["vendor_name"], "net": row["net"], "to": to, "cc": cc,
                "mode": "TEST" if test_to else "LIVE", "status": status, "error": error,
                "pdf": str(pdf),
            })
            progress.progress(i / len(selected), text=f"{i}/{len(selected)} · {row['vendor_name']} · {status}")
            if delay and i < len(selected):
                time.sleep(delay)

    append_log(results)
    res = pd.DataFrame(results)
    sent, failed = int((res["status"] == "Sent").sum()), int((res["status"] == "Failed").sum())
    (st.success if not failed else st.warning)(f"Done: {sent} sent, {failed} failed. Letters saved in {run_dir}")
    st.dataframe(res, hide_index=True, use_container_width=True)
    st.download_button("⬇ Send log (Excel)", to_excel(res), f"send_log_{run_dir.name}.xlsx")

if LOG_PATH.exists():
    with st.expander("Previous send log"):
        st.dataframe(pd.read_csv(LOG_PATH), hide_index=True, use_container_width=True)
