"""Session-scoped reporting dispositions; source records are never mutated."""
import json
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
                    history.append(decision(sheet,row,choice,reviewer,reason,owner,due,workbook_hash,context))
                except ValueError as err:
                    st.error(str(err))
                else:
                    st.rerun()
