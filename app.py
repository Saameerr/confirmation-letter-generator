"""Account Confirmation Letter app.

Run with:  streamlit run app.py
Company details are NOT stored in the code: they are entered once on the
Settings page and saved locally in data/settings.json (kept out of Git).
"""
import base64
import datetime as dt
import io
import json
import re
import uuid
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image as PILImage

from letters import build_pdf, calc, money, prepared_image, status_words

try:  # optional: automatic AD -> BS date conversion
    import nepali_datetime
except Exception:
    nepali_datetime = None

BASE = Path(__file__).parent
DATA_DIR = BASE / "data"
ASSET_DIR = BASE / "assets"
DATA_FILE = DATA_DIR / "parties.xlsx"
SETTINGS_FILE = DATA_DIR / "settings.json"
STAMP_FILE = ASSET_DIR / "stamp.png"
SIGN_FILE = ASSET_DIR / "signature.png"
DATA_DIR.mkdir(exist_ok=True)
ASSET_DIR.mkdir(exist_ok=True)

NATURES = ["Purchase", "Sales", "Purchase and Sales"]
OPEN_TYPES = ["Payable by us", "Receivable by us"]
PAGES = [
    ("Dashboard", ":material/dashboard:"),
    ("Data Entry", ":material/edit_note:"),
    ("Letters", ":material/description:"),
    ("Settings", ":material/settings:"),
]

# Neutral defaults only. No business details live in the source code.
DEFAULT_SETTINGS = {
    "name": "", "address": "", "phone": "", "pan": "", "email": "",
    "fy": "", "asat": "", "use_stamp": True,
    "stamp_h": 2.8, "sign_h": 2.2, "overlap": 35,
}
REQUIRED_SETTINGS = ["name", "address", "phone", "pan", "fy", "asat"]

# (record key, Excel heading, type: s = text, n = number)
FIELDS = [
    ("id", "ID", "s"), ("party", "Party Name", "s"), ("address", "Address", "s"),
    ("pan", "PAN No.", "s"), ("nature", "Nature", "s"), ("date_ad", "Date (AD)", "s"),
    ("date_bs", "Miti (BS)", "s"), ("fy", "Fiscal Year", "s"), ("asat", "As At", "s"),
    ("ref", "Ref No.", "s"), ("open_type", "Opening Type", "s"),
    ("open_amt", "Opening Amount", "n"), ("p_tax", "Purchase Taxable", "n"),
    ("p_exe", "Purchase Tax-exempt", "n"), ("p_vat", "Purchase VAT", "n"),
    ("s_tax", "Sales Taxable", "n"), ("s_exe", "Sales Tax-exempt", "n"),
    ("s_vat", "Sales VAT", "n"), ("paid", "Payment Made by Us", "n"),
    ("recv", "Payment Received by Us", "n"),
]


# ----------------------------------------------------------------- settings
def load_settings():
    s = dict(DEFAULT_SETTINGS)
    if SETTINGS_FILE.exists():
        try:
            s.update(json.loads(SETTINGS_FILE.read_text(encoding="utf-8")))
        except Exception:
            pass
    return s


def save_settings(s):
    SETTINGS_FILE.write_text(json.dumps(s, indent=2), encoding="utf-8")


def is_configured(s):
    return all(str(s.get(k, "")).strip() for k in REQUIRED_SETTINGS)


# ------------------------------------------------------------------ storage
def _num(x):
    try:
        return float(str(x).replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


def load_records():
    if not DATA_FILE.exists():
        return []
    df = pd.read_excel(DATA_FILE, dtype=str, keep_default_na=False)
    recs = []
    for _, row in df.iterrows():
        r = {}
        for key, head, typ in FIELDS:
            v = row.get(head, "")
            r[key] = _num(v) if typ == "n" else str(v).strip()
        if r["party"] and r["id"]:
            if r["nature"] not in NATURES:
                r["nature"] = NATURES[0]
            if r["open_type"] not in OPEN_TYPES:
                r["open_type"] = OPEN_TYPES[0]
            recs.append(r)
    return recs


def save_records(recs):
    """Write all records to the Excel file. Returns True on success."""
    rows = []
    for r in recs:
        c = calc(r)
        row = {head: r[key] for key, head, _ in FIELDS}
        row["Total Purchase"] = c["p_total"]
        row["Total Sales"] = c["s_total"]
        row["Closing (+ payable, - receivable)"] = c["closing"]
        row["Closing Status"] = status_words(c["closing"])[0]
        rows.append(row)
    cols = [h for _, h, _ in FIELDS] + ["Total Purchase", "Total Sales",
                                        "Closing (+ payable, - receivable)", "Closing Status"]
    df = pd.DataFrame(rows, columns=cols)
    try:
        with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as xw:
            df.to_excel(xw, index=False, sheet_name="Parties")
            ws = xw.sheets["Parties"]
            for i, col in enumerate(df.columns, 1):
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = max(14, min(32, len(col) + 4))
        return True
    except PermissionError:
        notify("error", "The data file is open in Excel. Close it and try again.")
        return False


# ------------------------------------------------------------ notifications
def notify(kind, text):
    """Queue a message; it is shown as a toast on the next render."""
    st.session_state["msg"] = (kind, text)


def flush_toast():
    m = st.session_state.pop("msg", None)
    if m:
        st.toast(m[1], icon="✅" if m[0] == "success" else "⚠️")


# --------------------------------------------------------------- form state
FORM_KEYS = ["f_date", "f_bs", "f_fy", "f_asat", "f_ref", "f_party", "f_address", "f_pan",
             "f_nature", "f_open_type", "f_open_amt", "f_p_tax", "f_p_exe", "f_p_vat",
             "f_s_tax", "f_s_exe", "f_s_vat", "f_paid", "f_recv"]


def bs_of(d):
    if nepali_datetime is None:
        return ""
    try:
        return nepali_datetime.date.from_datetime_date(d).strftime("%Y-%m-%d")
    except Exception:
        return ""


def reset_form():
    s = load_settings()
    today = dt.date.today()
    st.session_state.update({
        "f_date": today, "f_bs": bs_of(today), "f_fy": s["fy"], "f_asat": s["asat"], "f_ref": "",
        "f_party": "", "f_address": "", "f_pan": "", "f_nature": NATURES[0],
        "f_open_type": OPEN_TYPES[0], "f_open_amt": 0.0,
        "f_p_tax": 0.0, "f_p_exe": 0.0, "f_p_vat": 0.0,
        "f_s_tax": 0.0, "f_s_exe": 0.0, "f_s_vat": 0.0, "f_paid": 0.0, "f_recv": 0.0,
        "edit_id": None,
    })


def on_date_change():
    st.session_state["f_bs"] = bs_of(st.session_state["f_date"])


def form_to_record(rec_id):
    g = st.session_state
    return {
        "id": rec_id, "party": g["f_party"].strip(), "address": g["f_address"].strip(),
        "pan": g["f_pan"].strip(), "nature": g["f_nature"],
        "date_ad": g["f_date"].strftime("%Y-%m-%d"), "date_bs": g["f_bs"].strip(),
        "fy": g["f_fy"].strip(), "asat": g["f_asat"].strip(), "ref": g["f_ref"].strip(),
        "open_type": g["f_open_type"], "open_amt": float(g["f_open_amt"]),
        "p_tax": float(g["f_p_tax"]), "p_exe": float(g["f_p_exe"]), "p_vat": float(g["f_p_vat"]),
        "s_tax": float(g["f_s_tax"]), "s_exe": float(g["f_s_exe"]), "s_vat": float(g["f_s_vat"]),
        "paid": float(g["f_paid"]), "recv": float(g["f_recv"]),
    }


def record_to_form(r):
    g = st.session_state
    try:
        g["f_date"] = dt.datetime.strptime(r["date_ad"], "%Y-%m-%d").date()
    except ValueError:
        g["f_date"] = dt.date.today()
    g.update({"f_bs": r["date_bs"], "f_fy": r["fy"], "f_asat": r["asat"], "f_ref": r["ref"],
              "f_party": r["party"], "f_address": r["address"], "f_pan": r["pan"],
              "f_nature": r["nature"], "f_open_type": r["open_type"], "f_open_amt": r["open_amt"],
              "f_p_tax": r["p_tax"], "f_p_exe": r["p_exe"], "f_p_vat": r["p_vat"],
              "f_s_tax": r["s_tax"], "f_s_exe": r["s_exe"], "f_s_vat": r["s_vat"],
              "f_paid": r["paid"], "f_recv": r["recv"], "edit_id": r["id"]})


# ---------------------------------------------------------------- callbacks
def goto(page):
    st.session_state["page"] = page


def cb_save():
    rec = form_to_record(st.session_state["edit_id"] or uuid.uuid4().hex[:10])
    if not rec["party"]:
        notify("error", "Please enter the party name.")
        return
    if not rec["fy"] or not rec["asat"]:
        notify("error", "Fiscal year and 'As at' cannot be empty.")
        return
    recs = load_records()
    editing = st.session_state["edit_id"]
    if editing:
        recs = [rec if r["id"] == editing else r for r in recs]
    else:
        recs.append(rec)
    if save_records(recs):
        notify("success", f"Saved: {rec['party']}")
        reset_form()


def cb_edit(rid):
    for r in load_records():
        if r["id"] == rid:
            record_to_form(r)
            goto("Data Entry")
            return


def cb_confirm_delete(rid):
    recs = [r for r in load_records() if r["id"] != rid]
    if save_records(recs):
        notify("success", "Letter deleted.")
        st.session_state["del_id"] = None
        if st.session_state.get("edit_id") == rid:
            reset_form()


def cb_toggle_preview(rid):
    st.session_state["prev_id"] = None if st.session_state.get("prev_id") == rid else rid


# ------------------------------------------------------------------ helpers
def safe_name(s):
    s = re.sub(r'[\\/:*?"<>|]+', "", s).strip().replace(" ", "_")
    return s or "party"


def stamp_paths(use, s=None):
    """[stamp, signature] paths to print (an entry is None when not uploaded)."""
    if not use:
        return []
    return [str(p) if p.exists() else None for p in (STAMP_FILE, SIGN_FILE)]


def show_pdf(data):
    try:
        st.pdf(data, height=800)
    except Exception:  # fallback when the PDF component is not installed
        b64 = base64.b64encode(data).decode()
        st.markdown(f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="800"></iframe>',
                    unsafe_allow_html=True)


def sample_record(s):
    return {"id": "sample", "party": "Sample Party", "address": "Sample Address", "pan": "123456789",
            "nature": "Purchase", "date_ad": dt.date.today().strftime("%Y-%m-%d"), "date_bs": bs_of(dt.date.today()),
            "fy": s["fy"], "asat": s["asat"], "ref": "", "open_type": OPEN_TYPES[0], "open_amt": 10000.0,
            "p_tax": 100000.0, "p_exe": 0.0, "p_vat": 0.0, "s_tax": 0.0, "s_exe": 0.0, "s_vat": 0.0,
            "paid": 50000.0, "recv": 0.0}


# --------------------------------------------------------------- pages
def page_dashboard():
    st.title("Dashboard")
    recs = load_records()
    if not recs:
        st.info("No letters yet. Add the first one to see the summary here.")
        st.button("Add the first letter", type="primary", on_click=goto, args=("Data Entry",))
        return

    fys = sorted({r["fy"] for r in recs}, reverse=True)
    fy = st.selectbox("Fiscal year", ["All"] + fys, key="dash_fy")
    view = [r for r in recs if fy == "All" or r["fy"] == fy]

    payable = receivable = 0.0
    for r in view:
        cl = calc(r)["closing"]
        if cl > 0:
            payable += cl
        else:
            receivable += -cl
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Letters", len(view))
    c2.metric("Total payable by us", f"Rs. {payable:,.2f}")
    c3.metric("Total receivable by us", f"Rs. {receivable:,.2f}")
    c4.metric("Net position", f"Rs. {abs(payable - receivable):,.2f}",
              "we owe more" if payable > receivable else ("we are owed more" if receivable > payable else "balanced"),
              delta_color="off")

    st.subheader("Letters by nature")
    cols = st.columns(3)
    for col, n in zip(cols, NATURES):
        col.metric(n, sum(1 for r in view if r["nature"] == n))

    rows = []
    for r in view:
        cl = calc(r)["closing"]
        rows.append({"Party": r["party"], "Nature": r["nature"], "Closing": abs(cl),
                     "Status": "Payable" if cl > 0.004 else ("Receivable" if cl < -0.004 else "Nil")})
    df = pd.DataFrame(rows)

    left, right = st.columns(2)
    with left:
        st.subheader("Largest payables")
        top = df[df["Status"] == "Payable"].sort_values("Closing", ascending=False).head(5)
        if len(top) > 0:
            st.dataframe(top[["Party", "Closing"]], hide_index=True, width="stretch")
        else:
            st.caption("None")
    with right:
        st.subheader("Largest receivables")
        top = df[df["Status"] == "Receivable"].sort_values("Closing", ascending=False).head(5)
        if len(top) > 0:
            st.dataframe(top[["Party", "Closing"]], hide_index=True, width="stretch")
        else:
            st.caption("None")

    st.subheader("Recently added")
    recent = df.tail(5).iloc[::-1]
    st.dataframe(recent, hide_index=True, width="stretch")

    b1, b2, _ = st.columns([1, 1, 4])
    b1.button("New entry", type="primary", on_click=goto, args=("Data Entry",), width="stretch")
    b2.button("Open letters", on_click=goto, args=("Letters",), width="stretch")


def page_entry():
    editing = st.session_state.get("edit_id")
    st.title(f"Edit: {st.session_state['f_party'] or 'Letter'}" if editing else "Data Entry")
    if editing:
        st.info("You are editing a saved letter. Make your changes and press Update.")

    st.subheader("Letter Details")
    c1, c2, c3, c4 = st.columns(4)
    c1.date_input("Date (AD)", key="f_date", on_change=on_date_change)
    c2.text_input("Miti (BS)", key="f_bs", help="Filled automatically from the AD date. You can edit it.")
    c3.text_input("Fiscal year", key="f_fy", help="For example 2082/83")
    c4.text_input("As at", key="f_asat", help="Printed as 'as at the ...' in the letter. For example: end of Ashad 2083")
    st.text_input("Ref No. (optional)", key="f_ref")

    st.subheader("Party")
    c1, c2, c3 = st.columns([2, 2, 1])
    c1.text_input("Party name", key="f_party")
    c2.text_input("Address", key="f_address")
    c3.text_input("PAN No.", key="f_pan")

    st.subheader("Nature of Transaction")
    st.radio("Nature", NATURES, key="f_nature", horizontal=True, label_visibility="collapsed")
    nature = st.session_state["f_nature"]
    has_p = nature in ("Purchase", "Purchase and Sales")
    has_s = nature in ("Sales", "Purchase and Sales")

    st.subheader("Opening Balance")
    c1, c2 = st.columns(2)
    c1.selectbox("Opening balance is", OPEN_TYPES, key="f_open_type",
                 help="Payable by us: we owe the party. Receivable by us: the party owes us.")
    c2.number_input("Opening amount (Rs.)", min_value=0.0, step=100.0, format="%.2f", key="f_open_amt")

    if has_p:
        st.subheader("Purchase")
        c1, c2, c3 = st.columns(3)
        c1.number_input("Taxable (Rs.)", min_value=0.0, step=100.0, format="%.2f", key="f_p_tax")
        c2.number_input("Tax-exempt (Rs.)", min_value=0.0, step=100.0, format="%.2f", key="f_p_exe")
        c3.number_input("VAT (Rs.)", min_value=0.0, step=100.0, format="%.2f", key="f_p_vat")
    if has_s:
        st.subheader("Sales")
        c1, c2, c3 = st.columns(3)
        c1.number_input("Taxable (Rs.) ", min_value=0.0, step=100.0, format="%.2f", key="f_s_tax")
        c2.number_input("Tax-exempt (Rs.) ", min_value=0.0, step=100.0, format="%.2f", key="f_s_exe")
        c3.number_input("VAT (Rs.) ", min_value=0.0, step=100.0, format="%.2f", key="f_s_vat")

    st.subheader("Payments")
    st.caption("Not printed in the letter. Used only to work out the closing balance.")
    c1, c2 = st.columns(2)
    if has_p:
        c1.number_input("Payment made by us to the party (Rs.)", min_value=0.0, step=100.0, format="%.2f", key="f_paid")
    if has_s:
        c2.number_input("Payment received by us from the party (Rs.)", min_value=0.0, step=100.0, format="%.2f", key="f_recv")

    c = calc(form_to_record("preview"))
    words, _ = status_words(c["closing"])
    st.markdown("---")
    m1, m2, m3 = st.columns(3)
    m1.metric("Total purchase", money(c["p_total"]) if has_p else "-")
    m2.metric("Total sales", money(c["s_total"]) if has_s else "-")
    m3.metric("Closing balance", f"Rs. {abs(c['closing']):,.2f}", words, delta_color="off")

    b1, b2, _ = st.columns([1, 1, 4])
    b1.button("Update" if editing else "Save", type="primary", on_click=cb_save, width="stretch")
    b2.button("Cancel" if editing else "Clear form", on_click=reset_form, width="stretch")


def page_letters():
    st.title("Letters")
    s = load_settings()
    recs = load_records()
    if not recs:
        st.info("No letters saved yet. Add data on the Data Entry page first.")
        return

    fys = sorted({r["fy"] for r in recs}, reverse=True)
    c1, c2, c3 = st.columns([1, 2, 2])
    fy = c1.selectbox("Fiscal year", ["All"] + fys)
    q = c2.text_input("Search party")
    use_stamp = c3.checkbox("Print stamp and signature on the letters", value=bool(s["use_stamp"]),
                            help="Uses the images uploaded on the Settings page. Turn it off to sign and stamp by hand.")

    view = [r for r in recs if (fy == "All" or r["fy"] == fy) and q.lower() in r["party"].lower()]
    st.caption(f"{len(view)} letter(s)")
    if not view:
        return
    paths = stamp_paths(use_stamp)

    d1, d2, _ = st.columns([1.3, 1.3, 3])
    day = dt.date.today().strftime("%Y%m%d")
    d1.download_button("Download all (one PDF)", data=build_pdf(view, s, paths),
                       file_name=f"All_Letters_{day}.pdf", mime="application/pdf", width="stretch")
    zbuf = io.BytesIO()
    used = {}
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
        for r in view:
            nm = safe_name(r["party"])
            used[nm] = used.get(nm, 0) + 1
            z.writestr(f"{nm}.pdf" if used[nm] == 1 else f"{nm}_{used[nm]}.pdf", build_pdf([r], s, paths))
    d2.download_button("Download all (ZIP)", data=zbuf.getvalue(), file_name=f"Letters_{day}.zip",
                       mime="application/zip", width="stretch")

    st.markdown("---")
    widths = [3, 2, 2, 1.4, 1, 1, 1, 1]
    for col, t in zip(st.columns(widths), ["Party", "Nature", "Closing", "Date", "", "", "", ""]):
        col.markdown(f"**{t}**")

    for r in view:
        c = calc(r)
        _, w = status_words(c["closing"])
        row = st.columns(widths)
        row[0].write(f"{r['party']}  \n:grey[{r['fy']}]")
        row[1].write(r["nature"])
        row[2].write(f"{abs(c['closing']):,.2f} {w}".strip() if w else "Nil")
        row[3].write(r["date_ad"])
        row[4].download_button("PDF", data=build_pdf([r], s, paths), file_name=f"{safe_name(r['party'])}.pdf",
                               mime="application/pdf", key=f"dl_{r['id']}")
        row[5].button("View", key=f"pv_{r['id']}", on_click=cb_toggle_preview, args=(r["id"],))
        row[6].button("Edit", key=f"ed_{r['id']}", on_click=cb_edit, args=(r["id"],))
        if st.session_state.get("del_id") == r["id"]:
            row[7].button("Sure?", key=f"yes_{r['id']}", on_click=cb_confirm_delete, args=(r["id"],), type="primary")
            row[7].button("No", key=f"no_{r['id']}", on_click=lambda: st.session_state.update(del_id=None))
        else:
            row[7].button("Delete", key=f"del_{r['id']}", on_click=lambda i=r["id"]: st.session_state.update(del_id=i))
        if st.session_state.get("prev_id") == r["id"]:
            show_pdf(build_pdf([r], s, paths))


def save_image(uploaded, dest):
    img = PILImage.open(uploaded).convert("RGBA")
    img.thumbnail((1600, 1600))
    img.save(dest, format="PNG")


def page_settings():
    st.title("Settings")
    s = load_settings()
    if not is_configured(s):
        st.warning("Fill in the required company details below and save them to start using the app.")

    st.subheader("Company Details")
    c1, c2 = st.columns(2)
    name = c1.text_input("Company name *", s["name"])
    address = c2.text_input("Address *", s["address"])
    c1, c2, c3 = st.columns(3)
    phone = c1.text_input("Phone *", s["phone"])
    pan = c2.text_input("PAN No. *", s["pan"])
    email = c3.text_input("Email", s["email"])

    st.subheader("Fiscal Year")
    c1, c2 = st.columns(2)
    fy = c1.text_input("Default fiscal year *", s["fy"], placeholder="e.g. 2082/83",
                       help="Pre-filled on every new entry. Change it when the new fiscal year starts.")
    asat = c2.text_input("Default 'as at' text *", s["asat"], placeholder="e.g. end of Ashad 2083",
                         help="Printed as 'as at the ...' in the letter.")

    st.subheader("Stamp and Signature")
    use_stamp = st.checkbox("Print stamp and signature on letters by default", value=bool(s["use_stamp"]))
    st.caption("Upload a PNG or JPG. A plain white background is removed and empty margins are trimmed "
               "automatically, so the stamp and signature print as large as they can.")
    c1, c2 = st.columns(2)
    for col, label, path in [(c1, "Stamp", STAMP_FILE), (c2, "Signature", SIGN_FILE)]:
        with col:
            st.markdown(f"**{label}**")
            if path.exists():
                st.image(prepared_image(path), width=170)
                if st.button(f"Remove {label.lower()}", key=f"rm_{path.stem}"):
                    path.unlink()
                    notify("success", f"{label} removed.")
                    st.rerun()
            up = st.file_uploader(f"Upload {label.lower()}", type=["png", "jpg", "jpeg"], key=f"up_{path.stem}")
            if up is not None and st.button(f"Save {label.lower()}", key=f"sv_{path.stem}"):
                try:
                    save_image(up, path)
                    notify("success", f"{label} saved.")
                    st.rerun()
                except Exception:
                    notify("error", "This image could not be read. Try another PNG or JPG.")
                    st.rerun()

    st.markdown("**Size and position**")
    c1, c2, c3 = st.columns(3)
    stamp_h = c1.slider("Stamp height (cm)", 1.5, 4.5, float(s["stamp_h"]), 0.1)
    sign_h = c2.slider("Signature height (cm)", 1.0, 4.5, float(s["sign_h"]), 0.1)
    overlap = c3.slider("Overlap (%)", 0, 80, int(s["overlap"]), 5,
                        help="How much the signature overlaps the stamp. 0 puts them side by side.")

    if st.button("Save settings", type="primary"):
        new = dict(s)
        new.update({"name": name.strip(), "address": address.strip(), "phone": phone.strip(),
                    "pan": pan.strip(), "email": email.strip(), "fy": fy.strip(), "asat": asat.strip(),
                    "use_stamp": use_stamp, "stamp_h": stamp_h, "sign_h": sign_h, "overlap": overlap})
        if not is_configured(new):
            notify("error", "Please fill in all fields marked with *.")
        else:
            first_time = not is_configured(s)
            save_settings(new)
            notify("success", "Settings saved.")
            if not st.session_state.get("edit_id"):  # new entries start from the new defaults
                st.session_state["f_fy"], st.session_state["f_asat"] = new["fy"], new["asat"]
            if first_time:
                goto("Dashboard")
        st.rerun()

    if is_configured(s) and (STAMP_FILE.exists() or SIGN_FILE.exists()):
        with st.expander("Preview the signature area (saved settings)"):
            st.caption("Press 'Save settings' first if you changed the sliders.")
            show_pdf(build_pdf([sample_record(s)], s, stamp_paths(True)))


# ------------------------------------------------------------------- main
NAV_CSS = """
<style>
section[data-testid="stSidebar"] { width: 290px !important; }
section[data-testid="stSidebar"] .stButton > button {
    justify-content: flex-start; padding: 0.75rem 1rem; border-radius: 0.6rem;
}
section[data-testid="stSidebar"] .stButton > button p { font-size: 1.15rem; font-weight: 600; }
</style>
"""


def main():
    st.set_page_config(page_title="Confirmation Letters", page_icon="📄", layout="wide")
    st.markdown(NAV_CSS, unsafe_allow_html=True)

    if "f_party" not in st.session_state:
        reset_form()
    # keep form values alive while the user visits other pages
    for k in FORM_KEYS + ["edit_id"]:
        if k in st.session_state:
            st.session_state[k] = st.session_state[k]
    st.session_state.setdefault("page", "Dashboard")

    settings = load_settings()
    if not is_configured(settings):
        st.session_state["page"] = "Settings"  # first run: company details are needed first

    with st.sidebar:
        st.markdown(f"## {settings['name'] or 'Confirmation Letters'}")
        for label, icon in PAGES:
            st.button(label, icon=icon, key=f"nav_{label}", width="stretch",
                      type="primary" if st.session_state["page"] == label else "secondary",
                      on_click=goto, args=(label,), disabled=not is_configured(settings) and label != "Settings")
        st.caption(f"Saved letters: {len(load_records())}")

    flush_toast()
    {"Dashboard": page_dashboard, "Data Entry": page_entry,
     "Letters": page_letters, "Settings": page_settings}[st.session_state["page"]]()


main()
