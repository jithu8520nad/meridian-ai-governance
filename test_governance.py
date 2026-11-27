import io, zipfile
import openpyxl
import pytest
from governance_engine import (
    REQUIRED_SHEETS, parse_workbook, compute_financials,
    compute_delivery_and_rag, run_bounded_agents, verify_review_gate, execute_safe_rollover
)
from exporters import build_full_reporting_pack_zip

def _make_oct_and_nov(nov_actual=295000.0, dec_fcst_nov=355000.0):
    def _build(cutoff, is_nov):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for s in REQUIRED_SHEETS + ["Skills", "Resource Risks"]:
            wb.create_sheet(s)
        ws = wb["Controls"]
        ws.append([None, "Controls"]); ws.append([None, f"FICTIONAL TRAINING DATA | Workbook {'v0.3' if is_nov else 'v0.2'}"]); ws.append([]); ws.append([])
        ws.append([None, "Field", "Value", "Rule"]); ws.append([None, "Schema", "BANK-AI-PMO-v1", "Rule"]); ws.append([None, "Simulation day", 2 if is_nov else 1, "Run"]); ws.append([None, "Simulation cutoff", cutoff, "Month 1; stale note"]); ws.append([None, "Baseline authority", "Draft / assumptions", "Draft"]); ws.append([None, "Last simulation run", "RUN-000002" if is_nov else "RUN-000001", "Run"])
        ws_bb = wb["Budget Baseline"]; ws_bb.append([None]*4); ws_bb.append([None, "CBS ID", "Category", "24m EUR"]); ws_bb.append([None, "CBS-01", "Delivery", 8500000.0]); ws_bb.append([None, "CBS-10", "Contingency", 1500000.0])
        ws_mf = wb["Monthly Finance"]; ws_mf.append([None]*4); ws_mf.append([None, "Month", "Baseline EUR", "Actual EUR", "Forecast EUR", "Variance EUR", "Evidence status"])
        if not is_nov:
            ws_mf.append([None, "2026-10", 266666.67, 250000.0, 250000.0, -16666.67, "Booked"]); ws_mf.append([None, "2026-11", 266666.67, None, 383333.34, "Not supplied", "Forecast"]); ws_mf.append([None, "2026-12", 9466666.66, None, 9466666.66, "Not supplied", "Future"])
        else:
            ws_mf.append([None, "2026-10", 266666.67, 250000.0, 250000.0, -16666.67, "Booked"]); ws_mf.append([None, "2026-11", 266666.67, nov_actual, nov_actual, nov_actual - 266666.67, "Booked"]); ws_mf.append([None, "2026-12", 9466666.66, None, 9850000.0 - nov_actual if abs(dec_fcst_nov - 355000.0) < 1 else dec_fcst_nov, "Not supplied", "Future"])
        ws_cl = wb["Cost Ledger"]; ws_cl.append([None]*4); ws_cl.append([None, "Transaction ID", "Month", "Amount EUR", "Type booking accrual commitment"]); ws_cl.append([None, "TXN-001", "2026-10", 250000.0, "Fictional booking"])
        if is_nov:
            ws_cl.append([None, "TXN-002", "2026-11", nov_actual, "Fictional booking"])
        ws_ms = wb["Milestones"]; ws_ms.append([None]*4); ws_ms.append([None, "ID", "Phase", "Planning finish", "Current forecast finish", "Owner", "Authority"]); ws_ms.append([None, "MS-01", "Discovery", "2026-12-31", "2026-12-31", "Daniel Weber", "Proposed roadmap; dates not charter-approved"]); ws_ms.append([None, "MS-02", "Foundations", "2027-03-31", "2027-03-31", "Leo Martin", "Proposed roadmap; dates not charter-approved"])
        ws_raid = wb["RAID"]; ws_raid.append([None]*4); ws_raid.append([None, "id", "type", "title", "owner", "status", "severity", "statement", "response", "target_date"]); ws_raid.append([None, "RAID-I06", "Issue", "Source permission mapping incomplete", "Priya Nair", "Open", "Critical", "4 deny-case failures; affected connection blocked.", "Repair entitlement logic and retest.", "2026-11-06"]); ws_raid.append([None, "RAID-R01", "Risk", "Cash benefits fail to materialise", "Nora Fischer", "Assessed", "Critical", "Unvalidated cash pools.", "Validate counterfactuals.", "2026-12-18"])
        ws_act = wb["Actions"]; ws_act.append([None]*4); ws_act.append([None, "ID", "Action", "Owner", "Due / milestone", "Status", "Source"]); ws_act.append([None, "ACT-002", "Repair isolated access failures; retest before any affected connection", "Priya Nair / Sofia Klein", "2026-12-04", "Open", "EVT-000003; RAID-I06"])
        ws_chg = wb["Changes"]; ws_chg.append([None]*4); ws_chg.append([None, "ID", "Type", "Title", "Owner", "Status", "Approval", "Example only", "Incremental delivery EUR"]); ws_chg.append([None, "CHG-001", "Scope", "Example", "Nora Fischer", "Draft", "Not approved", "YES — example", 180000.0]); ws_chg.append([None, "CHG-005", "Forecast cost pressure", "Assess €100k delivery forecast uplift", "Daniel Weber", "Under assessment", "Not approved", "NO — live", 100000.0])
        ws_chgd = wb["Change Details"]; ws_chgd.append([None]*4); ws_chgd.append([None, "id", "title"]); ws_chgd.append([None, "CHG-001", "Example"])
        ws_ben = wb["Benefits Register"]; ws_ben.append([None]*4); ws_ben.append([None, "ID", "Benefit"]); ws_ben.append([None, "BEN-01", "Cash conversion"])
        ws_obs = wb["Benefit Observations"]; ws_obs.append([None]*4); ws_obs.append([None, "Observation ID", "Cash release EUR", "Avoided spend EUR"])
        ws_res = wb["Resources"]; ws_res.append([None]*4); ws_res.append([None, "ID", "Name"]); ws_res.append([None, "ST-01", "Maya Shah"])
        ws_rc = wb["Resource Calc"]; ws_rc.append([None]*4); ws_rc.append([None, "ID", "Gross FTE months", "Loaded effort EUR"]); ws_rc.append([None, "ST-01", 10.0, 100000.0])
        for s, cols in [("Sources", ["Source ID"]), ("Programme Baseline", ["Baseline element"]), ("Decisions", ["Decision ID"]), ("Recommendations", ["ID"]), ("Conflicts", ["ID", "Status"]), ("Event Log", ["Event ID"]), ("Update Log", ["Run ID", "Run key"]), ("Snapshots", ["Snapshot ID"]), ("Progress", ["Update ID"]), ("Backlog", ["ID", "Refinement status"])]:
            w = wb[s]; w.append([None]*4); w.append([None] + cols)
        ws_chk = wb["Checks"]; ws_chk.append([None]*4); ws_chk.append([None, "Control", "Calculated result", "Expected", "Status"]); ws_chk.append([None, "Budget total", 10000000, 10000000, "PASS"])
        b = io.BytesIO(); wb.save(b); return b.getvalue()
    return _build("2026-10-31", False), _build("2026-11-30", True)

def test_oct_nov_acceptance_criteria():
    oct_b, nov_b = _make_oct_and_nov()
    p_sheets, _, _ = parse_workbook(oct_b)
    c_sheets, c_ctrl, rep = parse_workbook(nov_b)
    assert rep.is_valid
    fin = compute_financials(c_sheets, c_ctrl, prev_sheets=p_sheets)
    assert fin["curr_month_actual_eur"] == 295000.0
    assert fin["prior_forecast_for_curr_month_eur"] == 383333.34
    assert fin["period_variance_eur"] == -88333.34
    assert fin["eac_eur"] == 10100000.0
    assert fin["prior_eac_eur"] == 10100000.0
    assert fin["rephased_only"] is True
    assert fin["live_crs"][0]["ID"] == "CHG-005"
    assert fin["live_crs"][0]["Approval"] == "Not approved"

    deliv = compute_delivery_and_rag(c_sheets, fin, rep)
    ms2 = deliv["unfinished_milestones"][:2]
    assert ms2[0]["ID"] == "MS-01" and ms2[0]["Plan"] == "2026-12-31" and ms2[0]["Owner"] == "Daniel Weber"
    assert ms2[1]["ID"] == "MS-02" and ms2[1]["Plan"] == "2027-03-31" and ms2[1]["Owner"] == "Leo Martin"
    assert "readiness unconfirmed" in ms2[0]["Combined Status"]
    assert "BLOCKED" in deliv["gate_readiness_rag"]
    assert "Priya Nair / Sofia Klein" in deliv["immediate_action"]
    assert "2026-12-04" in deliv["immediate_action"]

def test_dynamic_recalculation_different_values():
    oct_b, nov_b = _make_oct_and_nov(nov_actual=410000.0, dec_fcst_nov=9600000.0)
    p_sheets, _, _ = parse_workbook(oct_b)
    c_sheets, c_ctrl, rep = parse_workbook(nov_b)
    fin = compute_financials(c_sheets, c_ctrl, prev_sheets=p_sheets)
    assert fin["curr_month_actual_eur"] == 410000.0
    assert fin["period_variance_eur"] == round(410000.0 - 383333.34, 2)
    assert fin["rephased_only"] is False

def test_rollover_and_zip_exports():
    _, nov_b = _make_oct_and_nov()
    rolled, receipt = execute_safe_rollover(nov_b, "2026-12-31", [{"authorized": True, "description": "Close Nov", "evidence_id": "EVD-1"}])
    assert receipt["next_cutoff"] == "2026-12-31"
    c_sheets, c_ctrl, rep = parse_workbook(nov_b)
    fin = compute_financials(c_sheets, c_ctrl)
    deliv = compute_delivery_and_rag(c_sheets, fin, rep)
    agents = run_bounded_agents(c_sheets, c_ctrl, fin, deliv, rep)
    gates = {k: verify_review_gate(k, v["content_hash"], "h", fin["cutoff"], "v2.1", None) for k, v in agents["sections"].items()}
    zbytes, _ = build_full_reporting_pack_zip({"schema": "BANK-AI-PMO-v1", "cutoff": fin["cutoff"], "run_id": "RUN-000002", "wb_hash": "h", "overall_review_label": "Draft"}, c_sheets, fin, deliv, agents, gates, rolled, "Programme_Data_2026-12-31.xlsx")
    with zipfile.ZipFile(io.BytesIO(zbytes)) as zf:
        assert len(zf.namelist()) == 9
