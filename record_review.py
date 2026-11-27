"""Persistent simulated reporting dispositions; source records are never mutated."""
import json
import os
import sqlite3
from pathlib import Path
from contextlib import closing
from datetime import datetime, timezone
from hashlib import sha256

ACCEPT = "Accept for reporting"
RETURN = "Needs correction"
REJECT = "Reject"
SECTIONS = {
    "Monthly Finance": ("Finance","PMO","Governance"),
    "Cost Ledger": ("Finance","PMO","Governance"),
    "Milestones": ("PMO","Governance"),
    "Changes": ("Finance","PMO","BA","Governance"),
    "Change Details": ("Finance","PMO","BA","Governance"),
    "Backlog": ("BA",),
    "RAID": ("Governance","PMO"),
    "Conflicts": ("Governance","PMO","Finance","BA"),
    "Benefits Register": ("BA","Governance"),
    "Benefit Observations": ("BA","Governance"),
    "Controls": ("Governance","PMO","Finance","BA"),
    "Checks": ("Governance","PMO","Finance","BA"),
    "Rollover": ("Governance",),
}

def record_key(sheet, record):
    return sha256(json.dumps([sheet, record], sort_keys=True, default=str).encode()).hexdigest()

def decision(sheet, record, outcome, reviewer, reason, owner, due, workbook_hash, context):
    if sheet not in SECTIONS or outcome not in (ACCEPT, RETURN, REJECT):
        raise ValueError("Select a supported source and review outcome.")
    if not reviewer.strip() or not reason.strip():
        raise ValueError("Reviewer and reason are required.")
    if outcome != ACCEPT and (not owner.strip() or due is None):
        raise ValueError("Assign a correction owner and due date.")
    return dict(sheet=sheet, record=record, record_hash=record_key(sheet,record),
                outcome=outcome, reviewer=reviewer.strip(), reason=reason.strip(),
                owner=owner.strip(), due=due.isoformat() if due else None,
                timestamp=datetime.now(timezone.utc).isoformat(), workbook_hash=workbook_hash,
                context=context, affected_sections=list(SECTIONS[sheet]))

def open_items(history, context):
    latest = {item["record_hash"]: item for item in history if item["context"] == context}
    return [item for item in latest.values() if item["outcome"] != ACCEPT]

def database_path():
    return Path(os.environ.get("MERIDIAN_REVIEW_DB", str(Path.home() / ".meridian" / "record_reviews.sqlite3")))

def connect_history():
    path = database_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path, timeout=15)
    conn.execute("CREATE TABLE IF NOT EXISTS review_events (id INTEGER PRIMARY KEY, payload TEXT NOT NULL)")
    return conn

def load_history():
    with closing(connect_history()) as conn:
        return [json.loads(row[0]) for row in conn.execute("SELECT payload FROM review_events ORDER BY id")]

def unresolved_items(history):
    latest = {}
    for item in history:
        if item.get("closes_exception"):
            latest.pop(item["closes_exception"], None)
        if item["outcome"] != ACCEPT:
            latest[item["record_hash"]] = item
    return list(latest.values())

def save_decision(item, closure=None):
    with closing(connect_history()) as conn, conn:
        conn.execute("BEGIN IMMEDIATE")
        history = [json.loads(row[0]) for row in conn.execute("SELECT payload FROM review_events ORDER BY id")]
        pending = {x["record_hash"]: x for x in unresolved_items(history)}
        if closure:
            if closure not in pending or item["outcome"] != ACCEPT:
                raise ValueError("Select an open exception for closure.")
            item = dict(item, closes_exception=closure)
        elif item["outcome"] == ACCEPT and item["record_hash"] in pending:
            raise ValueError("Use the shared exception queue to document closure.")
        conn.execute("INSERT INTO review_events(payload) VALUES (?)", (json.dumps(item),))

def render_closure(st, exceptions, workbook_hash, context):
    if not exceptions:
        return
    with st.expander("Resolve an existing exception"):
        selected = st.selectbox("Open exception", range(len(exceptions)),
            format_func=lambda i: exceptions[i]["sheet"] + " | " + exceptions[i]["reason"])
        item = exceptions[selected]
        with st.form("exception_closure"):
            reviewer = st.text_input("Closure reviewer (simulated)")
            evidence = st.text_area("Correction evidence or documented override reference")
            confirmed = st.checkbox("I confirm the correction or authorised override has been verified")
            if st.form_submit_button("Close exception"):
                try:
                    if not confirmed or not evidence.strip():
                        raise ValueError("Verification and evidence are required.")
                    event = decision(item["sheet"], item["record"], ACCEPT, reviewer,
                                     evidence, "", None, workbook_hash, context)
                    save_decision(event, closure=item["record_hash"])
                except (ValueError, sqlite3.Error, OSError) as err:
                    st.error(str(err))
                else:
                    st.rerun()

def render_review(st, sources, history, workbook_hash, context, key):
    with st.expander("Review a source record"):
        st.caption("Simulated disposition only. Corrections must be applied by the data owner in a new workbook version.")
        sources={k:v for k,v in sources.items() if v}
        if not sources:
            st.caption("No records available.")
            return
        sheet=st.selectbox("Source", list(sources),key=key+"_sheet")
        records=sources[sheet]
        def label(i):
            r=records[i]
            sample=[str(v) for v in list(r.values())[:3] if v is not None]
            return f"{i+1} · " + " | ".join(sample)[:150]
        idx=st.selectbox("Record", range(len(records)),format_func=label,key=key+"_row_"+sheet)
        row=records[idx]
        st.json(row, expanded=False)
        with st.form(key+"_form_"+record_key(sheet,row)[:12]):
            choice=st.selectbox("Decision",[ACCEPT, RETURN, REJECT],index=None)
            reviewer=st.text_input("Simulated reviewer")
            reason=st.text_area("Reason / evidence reference")
            owner=st.text_input("Correction owner")
            due=st.date_input("Correction due",value=None)
            if st.form_submit_button("Record decision"):
                try:
                    save_decision(decision(sheet,row,choice,reviewer,reason,owner,due,workbook_hash,context))
                except (ValueError, sqlite3.Error, OSError) as err:
                    st.error(str(err))
                else:
                    st.rerun()
