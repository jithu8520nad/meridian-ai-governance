import os
import json
from record_review import render_review, open_items
from datetime import datetime, timezone
import pandas as pd
import streamlit as st
from governance_engine import (
    SCHEMA_VERSION, sha256_bytes, parse_workbook, compute_financials,
    compute_delivery_and_rag, run_bounded_agents, verify_review_gate, execute_safe_rollover
)
from exporters import build_full_reporting_pack_zip

st.set_page_config(page_title="Meridian Retail Bank · Executive Review & SteerCo", layout="wide")
st.markdown("""<style>
.block-container {padding-top: 1.2rem; padding-bottom: 1.5rem; max-width: 1180px;}
.bank-header {background: #0f2438; color: #ffffff; padding: 18px 22px; border-radius: 6px; margin-bottom: 12px;}
.bank-eyebrow {font-size: 11px; letter-spacing: .12em; text-transform: uppercase; color: #b8cde0; font-weight: 700;}
.bank-title {font-size: 23px; font-weight: 700; margin: 4px 0; color: #ffffff;}
.bank-sub {font-size: 13px; color: #d8e4ef;}
.status-strip {background: #ffffff; border: 1px solid #d5dde5; border-left: 6px solid #b45309; border-radius: 6px; padding: 12px 16px; margin-bottom: 14px; color: #16212c;}
.badge-AMBER {background: #fef3c7; color: #92400e; padding: 2px 8px; border-radius: 4px; font-weight: 700; font-size: 12px;}
.badge-RED {background: #fee2e2; color: #991b1b; padding: 2px 8px; border-radius: 4px; font-weight: 700; font-size: 12px;}
.badge-GREEN {background: #dcfce7; color: #166534; padding: 2px 8px; border-radius: 4px; font-weight: 700; font-size: 12px;}
.kpi-card {background: #ffffff; border: 1px solid #d5dde5; border-radius: 6px; padding: 13px 15px; color: #16212c;}
.kpi-lbl {font-size: 11px; text-transform: uppercase; letter-spacing: .06em; color: #5b6b7c; font-weight: 700;}
.kpi-val {font-size: 22px; font-weight: 700; margin: 4px 0; color: #0f2438;}
.kpi-sub {font-size: 12px; color: #5b6b7c;}
.var-pos {color: #166534; font-weight: 600;}
.var-neg {color: #991b1b; font-weight: 600;}
</style>""", unsafe_allow_html=True)

CONFIG_VERSION = "BANK-AI-PMO-APP-v2.2"

with st.sidebar:
    st.header("1. Controlled Master Workbooks")
    curr_up = st.file_uploader("Current Period Workbook (.xlsx) *", type=["xlsx"], key="curr")
    prev_up = st.file_uploader("Comparison / Prior Workbook (.xlsx)", type=["xlsx"], key="prev")
    if "run_prev_cmp" not in st.session_state:
        st.session_state["run_prev_cmp"] = True
    c_btn1, c_btn2 = st.columns(2)
    if c_btn1.button("Compare Prior File", type="primary", width="stretch"):
        st.session_state["run_prev_cmp"] = True
    if c_btn2.button("Single-File View", width="stretch"):
        st.session_state["run_prev_cmp"] = False
    st.divider()
    st.header("2. Optional LLM Commentary")
    api_key = st.text_input("Gemini API Key (blank = Deterministic)", type="password")
    st.divider()
    st.header("3. Next-Month Rollover Controls")
    next_cutoff = st.text_input("Next Month Cutoff (YYYY-MM-DD)", "2026-12-31")
    ev_auth = st.checkbox("Authorise staged event for rollover", value=False)
    ev_desc = st.text_input("Event description", "")
    ev_evid = st.text_input("Evidence ID (required if event supplied)", "")
    ev_rev = st.text_input("Simulated Reviewer Name", "Jithu Nadviledath")

curr_bytes = curr_up.getvalue() if curr_up else None
curr_name = curr_up.name if curr_up else None
prev_bytes = prev_up.getvalue() if prev_up else None
prev_name = prev_up.name if prev_up else None

if curr_bytes is None:
    xlsx_list = sorted([f for f in os.listdir(".") if f.endswith(".xlsx") and "Programme_Data" in f])
    nov_c = [f for f in xlsx_list if "November" in f or "v0.3" in f]
    oct_c = [f for f in xlsx_list if "October" in f or "v0.2" in f]
    pick_c = nov_c[-1] if nov_c else (xlsx_list[-1] if xlsx_list else None)
    if pick_c:
        with open(pick_c, "rb") as fh:
            curr_bytes, curr_name = fh.read(), pick_c
    if prev_bytes is None and oct_c:
        with open(oct_c[-1], "rb") as fh:
            prev_bytes, prev_name = fh.read(), oct_c[-1]

if curr_bytes is None:
    st.markdown("<div class='bank-header'><div class='bank-eyebrow'>Meridian Retail Bank · Governed AI Transformation</div><div class='bank-title'>Executive Review &amp; Steering Committee Portal</div></div>", unsafe_allow_html=True)
    st.warning("Upload a valid BANK-AI-PMO-v1 Programme_Data.xlsx workbook in the left sidebar to begin.")
    st.stop()

wb_hash = sha256_bytes(curr_bytes)
sheets, controls, val_rep = parse_workbook(curr_bytes)
if not val_rep.is_valid:
    st.error("BLOCKING SCHEMA / RECONCILIATION ERRORS — Finalisation halted:")
    for err in val_rep.blocking_errors:
        st.write(f"- {err}")
    st.stop()

prev_sheets, prev_controls = None, None
if prev_bytes is not None and st.session_state.get("run_prev_cmp", True):
    ps, pc, pr = parse_workbook(prev_bytes)
    if not pr.is_valid:
        st.error(f"PRIOR WORKBOOK VALIDATION FAILED ? Comparison stopped: {'; '.join(pr.blocking_errors)}. Fix the file or click Single-File View in the sidebar.")
    if pc.get("Simulation cutoff", "") >= controls.get("Simulation cutoff", ""):
        st.error(f"COMPARISON ORDER ERROR ? Prior cutoff ({pc.get('Simulation cutoff')}) must be earlier than current cutoff ({controls.get('Simulation cutoff')}). Comparison stopped.")
        st.stop()
    prev_sheets, prev_controls = ps, pc

fin = compute_financials(sheets, controls, prev_sheets=prev_sheets)
deliv = compute_delivery_and_rag(sheets, fin, val_rep)
agent_out = run_bounded_agents(sheets, controls, fin, deliv, val_rep, api_key=api_key)

if "reviews" not in st.session_state:
    st.session_state["reviews"] = {}
review_gates = {k: verify_review_gate(k, v["content_hash"], wb_hash, fin["cutoff"], CONFIG_VERSION, st.session_state["reviews"].get(k)) for k, v in agent_out["sections"].items()}
context = sha256_bytes(json.dumps([wb_hash, sha256_bytes(prev_bytes) if prev_sheets is not None else None,
                                   fin["cutoff"], CONFIG_VERSION]).encode())
history = st.session_state.setdefault("record_review_history", [])
exceptions = open_items(history, context)
for item in exceptions:
    for section in item["affected_sections"]:
        if section in review_gates:
            st.session_state["reviews"].pop(section, None)
            review_gates[section] = {"accepted": False, "status": "BLOCKED BY RECORD EXCEPTION",
                                     "reason": "Resolve or correct disputed source data before sign-off."}
if exceptions:
    st.warning(f"{len(exceptions)} open record exception(s). Affected figures are provisional; exports are held.")
all_accepted = all(g["accepted"] for g in review_gates.values())

def review_panel(key, *names, extra=None):
    sources = {name: json.loads(sheets[name].to_json(orient="records", date_format="iso"))
               for name in names if name in sheets}
    sources.update(extra or {})
    render_review(st, sources, history, wb_hash, context, key)

overall_review_label = "Simulated Owner-Acknowledged Pack" if all_accepted else "Draft · Pending Section Review"
comp_label = f"{prev_controls.get('Simulation cutoff')} ({prev_name})" if prev_controls else "Current workbook baseline (no prior file uploaded)"

st.markdown(f"<div class='bank-header'><div class='bank-eyebrow'>Meridian Retail Bank · Governed AI Transformation · {overall_review_label}</div><div class='bank-title'>Executive Review &amp; Steering Committee Pack</div><div class='bank-sub'>Master Source: {curr_name} &nbsp;|&nbsp; Cutoff: {fin['cutoff']} &nbsp;|&nbsp; Run: {controls.get('Last simulation run')} &nbsp;|&nbsp; Envelope Authority: {fin['authority_status']}</div></div>", unsafe_allow_html=True)

tab_exec, tab_rev, tab_cost, tab_ms, tab_raid, tab_roll = st.tabs([
    "1. Executive Review",
    "2. Five-Role Workflow & Review Gate",
    "3. Cost & Variance Ledger",
    "4. Milestones & Change Control",
    "5. RAID, Conflicts & Benefits",
    "6. Next-Month Rollover & Exports"
])

with tab_exec:
    perf_cls = "badge-GREEN" if deliv["performance_rag"] == "GREEN" else ("badge-RED" if deliv["performance_rag"] == "RED" else "badge-AMBER")
    gate_cls = "badge-GREEN" if "GREEN" in deliv["gate_readiness_rag"] else "badge-RED"
    st.markdown(f"<div class='status-strip'><div style='display:flex; flex-wrap:wrap; gap:18px; justify-content:space-between; margin-bottom:8px;'><div><strong>Reporting Month:</strong> {fin['period_ym']} (Cutoff {fin['cutoff']})</div><div><strong>Comparison Baseline:</strong> {comp_label}</div><div><strong>Overall Programme RAG:</strong> <span class='{perf_cls}'>{deliv['performance_rag']}</span></div><div><strong>Release Readiness:</strong> <span class='{gate_cls}'>{deliv['gate_readiness_rag']}</span></div></div><div style='margin-bottom:4px;'><strong>Primary Driver:</strong> {deliv['primary_driver']}</div><div style='margin-bottom:4px;'><strong>Immediate Action:</strong> {deliv['immediate_action']}</div><div style='font-size:11.5px; color:#5b6b7c;'><em>RAG Rule Basis: {deliv['rag_rule_basis']}</em></div></div>", unsafe_allow_html=True)

    prior_fcst = fin["prior_forecast_for_curr_month_eur"]
    prior_fcst_disp = f"€{prior_fcst:,.2f}" if prior_fcst is not None else f"€{fin['curr_month_baseline_eur']:,.2f} (Baseline)"
    var_eur, var_pct = fin["period_variance_eur"], fin["period_variance_pct"]
    var_dir = "below forecast" if var_eur < 0 else ("above forecast" if var_eur > 0 else "on forecast")
    var_cls = "var-pos" if var_eur <= 0 else "var-neg"
    prior_eac_txt = f"Unchanged vs prior EAC (€{fin['prior_eac_eur']/1e6:.2f}M)" if (fin["prior_eac_eur"] is not None and abs(fin["eac_eur"] - fin["prior_eac_eur"]) < 0.01) else (f"Prior EAC: €{fin['prior_eac_eur']/1e6:.2f}M" if fin["prior_eac_eur"] is not None else "Single-period view")

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"<div class='kpi-card'><div class='kpi-lbl'>Current-Month Actual ({fin['period_ym']})</div><div class='kpi-val'>€{fin['curr_month_actual_eur']:,.2f}</div><div class='kpi-sub'>Reconciled to Cost Ledger · Cum. €{fin['cum_actual_eur']:,.0f}</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='kpi-card'><div class='kpi-lbl'>Prior Forecast for {fin['period_ym']}</div><div class='kpi-val'>{prior_fcst_disp}</div><div class='kpi-sub'>Plan baseline: €{fin['curr_month_baseline_eur']:,.2f}</div></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='kpi-card'><div class='kpi-lbl'>Period Variance ({fin['period_ym']})</div><div class='kpi-val {var_cls}'>€{abs(var_eur):,.2f} {var_dir}</div><div class='kpi-sub'>{var_pct:+.2f}% vs prior forecast</div></div>", unsafe_allow_html=True)
    c4.markdown(f"<div class='kpi-card'><div class='kpi-lbl'>Expected Final Cost (EAC)</div><div class='kpi-val'>€{fin['eac_eur']/1e6:.2f}M</div><div class='kpi-sub'>{prior_eac_txt} · €{fin['eac_vs_envelope_eur']:+,.0f} vs €{fin['envelope_eur']/1e6:.1f}M envelope ({fin['authority_status']})</div></div>", unsafe_allow_html=True)

    st.write("")
    v_left, v_right = st.columns([1.45, 1.0])
    with v_left:
        st.subheader("3. Explain the Variance & Change Control Position")
        if fin["rephased_only"]:
            what_moved = f"€{abs(var_eur):,.2f} ({abs(var_pct):.2f}%) lower actual spend in {fin['period_ym']} (€{fin['curr_month_actual_eur']:,.2f} vs €{prior_fcst:,.2f} prior forecast), matched by an equal +€{abs(var_eur):,.2f} increase in subsequent monthly forecast."
            why_moved = "Spend moved to later months; underlying cause and recurrence unconfirmed. Do not treat rephasing as a saving."
            recurrence_txt = "Unconfirmed (rephased across periods; total EAC unchanged at €10.10M)."
        else:
            what_moved = f"€{abs(var_eur):,.2f} ({var_pct:+.2f}%) variance in {fin['period_ym']} against comparison forecast."
            why_moved = "Variance reflects actual ledger bookings vs prior forecast; underlying operational driver requires cost-owner confirmation."
            recurrence_txt = "Unconfirmed — awaiting cost-owner classification."
        st.markdown(f"- **1. What moved:** {what_moved}\n- **2. Why it moved:** {why_moved}\n- **3. Will it recur:** {recurrence_txt}\n- **4. What Finance should do next:** Confirm subsequent monthly phasing with Elena Rossi and Leo Martin; escalate €100k whole-programme EAC pressure (CHG-005) to F02/F01.")
        if fin["live_crs"]:
            st.markdown("**Existing Live Change Requests (No Automatic Creation or Approval):**")
            st.dataframe(pd.DataFrame(fin["live_crs"])[["ID", "Type", "Title", "Owner", "Incremental delivery EUR", "Status", "Approval"]], width="stretch", hide_index=True)

    with v_right:
        st.subheader("Cost Owner Variance Note (Draft)")
        with st.form("cost_owner_note_form"):
            co_cause = st.text_input("Cause", value="Supplier/platform onboarding rephased to next month")
            co_rec = st.selectbox("Recurrence Classification", ["Unconfirmed", "One-off", "Recurring"])
            co_evid = st.text_input("Evidence Reference", value="SIM-COST-02 / Monthly Finance")
            co_act = st.text_input("Proposed Finance Action", value="Monitor next-month actuals before releasing deferred ETC")
            if st.form_submit_button("Save Cost Owner Note (Draft)"):
                st.session_state["cost_owner_note"] = {"cause": co_cause, "recurrence": co_rec, "evidence": co_evid, "action": co_act}
        saved_note = st.session_state.get("cost_owner_note")
        if saved_note:
            st.warning(f"**DRAFT NOTE (Not Formally Reviewed):** Cause: {saved_note['cause']} | Recurrence: **{saved_note['recurrence']}** | Evidence: `{saved_note['evidence']}` | Action: {saved_note['action']}")
            if saved_note["recurrence"] == "Recurring":
                st.error("Recurring Impact Assessment: Alters remaining ETC and EAC beyond Daniel Weber's €100k monthly reallocation threshold (DR03); formal F02/F01 change governance required.")

    st.divider()
    st.subheader("4. Immediate Milestones (Next Two Unfinished Gates)")
    ms_cols = ["ID", "Phase", "Plan", "Forecast", "Combined Status", "Blocker", "Next Action", "Owner", "Due"]
    st.dataframe(pd.DataFrame(deliv["unfinished_milestones"][:2], columns=ms_cols), width="stretch", hide_index=True)

    st.subheader("5. Priority Risks & Issues (Top 3 Exceptions)")
    if deliv["overdue_actions"]:
        st.caption(f"Overdue Open Actions as of {fin['cutoff']}: " + ", ".join([f"{a['ID']} ({a['Owner']}, due {a['Due']})" for a in deliv["overdue_actions"]]))
    exc_cols = ["Record Type", "IDs", "Severity", "Title & Driving Evidence", "Impact", "Action", "Owner", "Due Date"]
    st.dataframe(pd.DataFrame(deliv["top_exceptions"], columns=exc_cols), width="stretch", hide_index=True)

with tab_rev:
    st.subheader("Five-Role Operating Model & Section-Level Human Review Gate")
    st.info("Orchestrator coordinates run & cutoff; Agent 0 is sole writer; PMO, BA and Governance are read-only. Reviewer entries are **simulated acknowledgements** bound to the workbook SHA-256 hash and cutoff.")
    for sec_key, sec_data in agent_out["sections"].items():
        g = review_gates[sec_key]
        with st.container(border=True):
            ca, cb = st.columns([2.2, 1.3])
            with ca:
                st.markdown(f"**{sec_data['title']}** (`{sec_data['mode']}`)")
                for b in sec_data["bullets"]:
                    st.write(f"- {b}")
                st.caption(f"Gate Status: **{g['status']}**")
            with cb:
                r_name = st.text_input("Simulated Reviewer", value="Owner Name", key=f"rev_{sec_key}")
                r_out = st.selectbox("Outcome", ["Pending", "Approved", "Changes Requested", "Rejected"], key=f"out_{sec_key}")
                r_com = st.text_input("Review Note", key=f"com_{sec_key}")
                if st.button("Record Simulated Review", key=f"btn_{sec_key}", disabled=any(sec_key in x["affected_sections"] for x in exceptions)):
                    st.session_state["reviews"][sec_key] = {"reviewer": f"{r_name} (Simulated Acknowledgement)", "outcome": r_out, "comments": r_com, "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"), "workbook_hash": wb_hash, "cutoff": fin["cutoff"], "section_hash": sec_data["content_hash"], "config_version": CONFIG_VERSION}
                    st.rerun()

with tab_cost:
    st.subheader("Monthly Finance & Cost Ledger Reconciliation")
    review_panel("cost", "Monthly Finance", "Cost Ledger")
    mf_disp = sheets["Monthly Finance"].copy()
    for col in ["Baseline EUR", "Actual EUR", "Forecast EUR", "Variance EUR"]:
        if col in mf_disp.columns:
            mf_disp[col] = pd.to_numeric(mf_disp[col], errors="coerce").round(2)
    if prev_sheets and "Monthly Finance" in prev_sheets:
        pmf_map = prev_sheets["Monthly Finance"].set_index("Month")["Forecast EUR"].apply(lambda x: round(float(x), 2) if pd.notna(x) else None).to_dict()
        for idx, row in mf_disp.iterrows():
            m_key = row["Month"]
            if pd.notna(row["Actual EUR"]) and m_key in pmf_map and pmf_map[m_key] is not None:
                mf_disp.at[idx, "Forecast EUR"] = pmf_map[m_key]
    mf_disp["Variance EUR"] = (mf_disp["Actual EUR"] - mf_disp["Forecast EUR"]).round(2)
    mf_disp["Variance %"] = ((mf_disp["Variance EUR"] / mf_disp["Forecast EUR"]) * 100.0).round(2)
    cols_order = [c for c in ["Month", "Baseline EUR", "Forecast EUR", "Actual EUR", "Variance EUR", "Variance %", "Evidence status"] if c in mf_disp.columns]
    st.dataframe(mf_disp[cols_order], width="stretch", height=280, hide_index=True)
    st.markdown("**Cost Ledger (Bookings, Accruals & Commitments):**")
    st.dataframe(sheets["Cost Ledger"], width="stretch", height=240, hide_index=True)

with tab_ms:
    st.subheader("Milestones, Change Control & Backlog")
    review_panel("delivery", "Milestones", "Changes", "Change Details", "Backlog")
    st.dataframe(sheets["Milestones"], width="stretch", height=180)
    st.dataframe(sheets["Changes"], width="stretch", height=200)
    st.dataframe(sheets["Backlog"], width="stretch", height=240)

with tab_raid:
    st.subheader("RAID Register, Source Conflicts, Benefits & Checks")
    st.dataframe(sheets["RAID"], width="stretch", height=220)
    review_panel("controls", "RAID", "Conflicts", "Benefits Register", "Benefit Observations",
                 extra={"Checks": val_rep.checks_summary,
                        "Controls": [{"Finding": warning} for warning in val_rep.warnings]})
    if val_rep.warnings:
        for w in val_rep.warnings:
            st.warning(w)
    st.dataframe(pd.DataFrame(val_rep.checks_summary), width="stretch", height=210)
    st.dataframe(sheets["Conflicts"], width="stretch", height=200)
    st.dataframe(sheets["Benefits Register"], width="stretch", height=220)

meta = {"schema": SCHEMA_VERSION, "cutoff": fin["cutoff"], "run_id": controls.get("Last simulation run", "RUN-UNKNOWN"), "wb_hash": wb_hash, "overall_review_label": overall_review_label}

with tab_roll:
    st.subheader("Safe Next-Month Rollover & Complete 10-Artefact Reporting Pack")
    st.markdown("### Shared exceptions queue")
    st.caption("Items are shared across tabs for this session. Download the history before leaving. A new workbook requires fresh review.")
    staged = {"Next cutoff": next_cutoff, "Event": ev_desc, "Evidence": ev_evid, "Staged": ev_auth}
    review_panel("rollover", extra={"Rollover": [staged]})
    if history:
        st.download_button("Download record review history", json.dumps(history, indent=2),
                           "Meridian_Record_Review_History.json", "application/json")
    if exceptions:
        st.dataframe([{ "Source": x["sheet"], "Record": str(x["record"]), "Decision": x["outcome"],
                        "Reason": x["reason"], "Owner": x["owner"], "Due": x["due"],
                        "Sections": ", ".join(x["affected_sections"])} for x in exceptions],
                     hide_index=True)
        st.warning("Reporting pack and rollover exports are held until the open exceptions are resolved.")
        st.stop()
    st.caption("No open reviewer exceptions for these inputs.")
    rollover_bytes, rollover_name, receipt = None, None, None
    staged_events = [{"authorized": ev_auth, "description": ev_desc, "evidence_id": ev_evid, "reviewer": ev_rev}] if ev_desc.strip() else []
    try:
        rollover_bytes, receipt = execute_safe_rollover(curr_bytes, next_cutoff, staged_events, CONFIG_VERSION)
        rollover_name = f"Programme_Data_{next_cutoff}.xlsx"
        st.success(f"Next-month rollover verified ({fin['cutoff']} -> {next_cutoff}) | Output SHA-256: `{receipt['output_sha256'][:16]}`")
    except Exception as e:
        st.error(f"Rollover blocked: {e}")

    zip_bytes, file_map = build_full_reporting_pack_zip(meta, sheets, fin, deliv, agent_out, review_gates, rollover_bytes, rollover_name)
    st.download_button("📦 Download Complete Reporting Pack (.zip)", zip_bytes, f"Meridian_Reporting_Pack_{fin['period_ym']}.zip", "application/zip", type="primary", width="stretch")

    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button("1. Monthly Status Report (.docx)", file_map[f"01_Monthly_Status_Report_{fin['period_ym']}.docx"], f"01_Monthly_Status_Report_{fin['period_ym']}.docx", width="stretch")
        st.download_button("2. Steering Committee Deck (.pptx)", file_map[f"02_SteerCo_Deck_{fin['period_ym']}.pptx"], f"02_SteerCo_Deck_{fin['period_ym']}.pptx", width="stretch")
        st.download_button("3. RAID Register (.xlsx)", file_map[f"03_RAID_Register_{fin['period_ym']}.xlsx"], f"03_RAID_Register_{fin['period_ym']}.xlsx", width="stretch")
    with d2:
        st.download_button("4. Cost Analysis Workbook (.xlsx)", file_map[f"04_Cost_Analysis_{fin['period_ym']}.xlsx"], f"04_Cost_Analysis_{fin['period_ym']}.xlsx", width="stretch")
        st.download_button("5. Benefits Analysis (.xlsx)", file_map[f"05_Benefits_Analysis_{fin['period_ym']}.xlsx"], f"05_Benefits_Analysis_{fin['period_ym']}.xlsx", width="stretch")
        st.download_button("6. Backlog Review (.xlsx)", file_map[f"06_Backlog_Review_{fin['period_ym']}.xlsx"], f"06_Backlog_Review_{fin['period_ym']}.xlsx", width="stretch")
    with d3:
        st.download_button("7. Monthly Change Summary (.docx)", file_map[f"07_Change_Summary_{fin['period_ym']}.docx"], f"07_Change_Summary_{fin['period_ym']}.docx", width="stretch")
        st.download_button("8. Five Role Outputs (.json)", file_map[f"08_Role_Outputs_{fin['period_ym']}.json"], f"08_Role_Outputs_{fin['period_ym']}.json", width="stretch")
        if rollover_bytes and rollover_name:
            st.download_button(f"9. Next-Month Workbook ({rollover_name})", rollover_bytes, rollover_name, width="stretch")
