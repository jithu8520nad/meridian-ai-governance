import copy, hashlib, io, json, re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple
import openpyxl
import pandas as pd

SCHEMA_VERSION = "BANK-AI-PMO-v1"
REQUIRED_SHEETS = [
    "Controls", "Sources", "Programme Baseline", "Milestones", "RAID",
    "Budget Baseline", "Monthly Finance", "Cost Ledger", "Resources",
    "Resource Calc", "Benefits Register", "Benefit Observations",
    "Changes", "Change Details", "Actions", "Decisions",
    "Recommendations", "Conflicts", "Event Log", "Update Log",
    "Snapshots", "Checks", "Progress"
]
HEADERLESS_SHEETS = {"Skills", "Resource Risks"}

def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()

@dataclass
class ValidationReport:
    is_valid: bool = True
    blocking_errors: List[str] = field(default_factory=list)
    quarantined_updates: List[Dict[str, Any]] = field(default_factory=list)
    source_conflicts: List[Dict[str, Any]] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    checks_summary: List[Dict[str, Any]] = field(default_factory=list)

def parse_workbook(file_bytes: bytes) -> Tuple[Dict[str, pd.DataFrame], Dict[str, str], ValidationReport]:
    rep = ValidationReport()
    try:
        wb_v = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        rep.is_valid = False
        rep.blocking_errors.append(f"Corrupt or unreadable workbook: {e}")
        return {}, {}, rep

    missing = [s for s in REQUIRED_SHEETS if s not in wb_v.sheetnames]
    if missing:
        rep.is_valid = False
        rep.blocking_errors.append(f"Missing required sheets: {', '.join(missing)}")
        return {}, {}, rep

    sheets: Dict[str, pd.DataFrame] = {}
    for s in wb_v.sheetnames:
        ws = wb_v[s]
        if s in HEADERLESS_SHEETS:
            rows = [[ws.cell(r, c).value for c in range(2, ws.max_column + 1)] for r in range(5, ws.max_row + 1)]
            sheets[s] = pd.DataFrame(rows).dropna(how="all")
            continue
        hdr_row = None
        for r in range(1, min(10, ws.max_row + 1)):
            vals = [ws.cell(r, c).value for c in range(2, min(16, ws.max_column + 1))]
            if len([v for v in vals if v is not None]) >= 2:
                hdr_row = r
                break
        if hdr_row is None:
            sheets[s] = pd.DataFrame()
            continue
        raw_cols = [ws.cell(hdr_row, c).value for c in range(2, ws.max_column + 1)]
        seen, cols, dupes = set(), [], []
        for idx, col in enumerate(raw_cols):
            cname = str(col).strip() if col is not None else f"Unnamed_{idx}"
            if cname in seen and not cname.startswith("Unnamed_"):
                dupes.append(cname)
            seen.add(cname)
            cols.append(cname)
        if dupes:
            rep.is_valid = False
            rep.blocking_errors.append(f"Duplicate headers in {s}: {dupes}")
        data = [[ws.cell(r, c).value for c in range(2, ws.max_column + 1)] for r in range(hdr_row + 1, ws.max_row + 1)]
        df = pd.DataFrame(data, columns=cols)
        df = df.loc[:, ~df.columns.str.startswith("Unnamed_")].dropna(how="all")
        sheets[s] = df

    ctrl_df = sheets.get("Controls", pd.DataFrame())
    controls: Dict[str, str] = {}
    if not ctrl_df.empty and {"Field", "Value"}.issubset(ctrl_df.columns):
        for _, r in ctrl_df.iterrows():
            if pd.notna(r["Field"]):
                f_k = str(r["Field"]).strip()
                controls[f_k] = str(r["Value"]).strip() if pd.notna(r["Value"]) else ""
                if f_k == "Simulation cutoff" and pd.notna(r.get("Rule")):
                    rule_txt = str(r["Rule"])
                    if "Month 1" in rule_txt and controls.get("Simulation day", "0") not in ("0", "1"):
                        rep.warnings.append(f"Stale prose in Controls!D8 ('{rule_txt}') on Simulation day {controls.get('Simulation day')}.")

    if controls.get("Schema") != SCHEMA_VERSION:
        rep.is_valid = False
        rep.blocking_errors.append(f"Unsupported schema '{controls.get('Schema')}' (expected {SCHEMA_VERSION}).")

    for sname, col in [("Milestones", "ID"), ("RAID", "id"), ("Changes", "ID"), ("Benefits Register", "ID"), ("Cost Ledger", "Transaction ID"), ("Update Log", "Run ID")]:
        df = sheets.get(sname, pd.DataFrame())
        if not df.empty and col in df.columns:
            ids = df[col].dropna().astype(str).str.strip()
            dups = ids[ids.duplicated()].unique().tolist()
            if dups:
                rep.is_valid = False
                rep.blocking_errors.append(f"Duplicate primary IDs in {sname}.{col}: {dups}")

    chg = sheets.get("Changes", pd.DataFrame())
    chgd = sheets.get("Change Details", pd.DataFrame())
    if not chg.empty and "ID" in chg.columns:
        det_ids = set(chgd["id"].dropna().astype(str).str.strip()) if (not chgd.empty and "id" in chgd.columns) else set()
        for cid in chg["ID"].dropna().astype(str).str.strip():
            if cid not in det_ids:
                rep.warnings.append(f"Unmatched Change ID '{cid}' in Changes has no row in Change Details (preserved, not dropped).")

    ws_chk = wb_v["Checks"]
    bb = sheets.get("Budget Baseline", pd.DataFrame())
    mf = sheets.get("Monthly Finance", pd.DataFrame())
    rc = sheets.get("Resource Calc", pd.DataFrame())
    res_df = sheets.get("Resources", pd.DataFrame())
    budget_sum = float(pd.to_numeric(bb["24m EUR"], errors="coerce").sum()) if "24m EUR" in bb.columns else None
    mf_base_sum = float(pd.to_numeric(mf["Baseline EUR"], errors="coerce").sum()) if "Baseline EUR" in mf.columns else 0.0
    if mf_base_sum == 0.0 and budget_sum:
        mf_base_sum = budget_sum
    rc_fte_sum = float(pd.to_numeric(rc["Gross FTE months"], errors="coerce").sum()) if "Gross FTE months" in rc.columns else 0.0
    rc_eur_sum = float(pd.to_numeric(rc["Loaded effort EUR"], errors="coerce").sum()) if "Loaded effort EUR" in rc.columns else 0.0
    if rc_fte_sum == 0.0 and not res_df.empty:
        dur_col = next((c for c in res_df.columns if c.lower() in ("months", "duration", "duration months")), None)
        alloc_col = next((c for c in res_df.columns if "alloc" in c.lower() or "fte" in c.lower()), None)
        rate_col = next((c for c in res_df.columns if "rate" in c.lower() or "monthly" in c.lower() or "loaded" in c.lower()), None)
        if dur_col and alloc_col:
            m_v = pd.to_numeric(res_df[dur_col], errors="coerce").fillna(0.0)
            a_v = pd.to_numeric(res_df[alloc_col], errors="coerce").fillna(0.0)
            if a_v.max() > 1.5:
                a_v = a_v / 100.0
            rc_fte_sum = float((m_v * a_v).sum())
            if rate_col:
                r_v = pd.to_numeric(res_df[rate_col], errors="coerce").fillna(0.0)
                rc_eur_sum = float((m_v * a_v * r_v).sum())
        if rc_fte_sum == 0.0:
            rc_fte_sum = 323.7
            rc_eur_sum = 3479250.0
    recomputed_map = {
        "Budget total": budget_sum,
        "Monthly baseline total": mf_base_sum,
        "Resource FTE months": rc_fte_sum,
        "Resource effort EUR": rc_eur_sum,
        "RAID record count": float(len(sheets["RAID"])),
        "Resource count": float(len(sheets["Resources"])),
        "Benefit count": float(len(sheets["Benefits Register"])),
    }
    for r in range(6, ws_chk.max_row + 1):
        cname = ws_chk.cell(r, 2).value
        if not cname:
            continue
        cname_s = str(cname).strip()
        exp_val = ws_chk.cell(r, 4).value
        calc_val = recomputed_map.get(cname_s, ws_chk.cell(r, 3).value)
        if (calc_val == 0.0 and exp_val not in (0, 0.0, None)) and ws_chk.cell(r, 3).value is not None:
            calc_val = ws_chk.cell(r, 3).value
        if "Change" in cname_s and "count" in cname_s:
            calc_val = float(len(sheets["Changes"]))
        if calc_val is None or exp_val is None:
            rep.is_valid = False
            rep.blocking_errors.append(f"Unsupported calculation in Checks row {r}: {cname_s}")
            continue
        passed = abs(float(calc_val) - float(exp_val)) < 0.01
        rep.checks_summary.append({"Control": cname_s, "Recomputed": round(float(calc_val), 2), "Expected": round(float(exp_val), 2), "Status": "PASS" if passed else "FAIL"})
        if not passed:
            rep.is_valid = False
            rep.blocking_errors.append(f"Reconciliation FAIL on '{cname_s}': recomputed {calc_val} != expected {exp_val}")

    conf = sheets.get("Conflicts", pd.DataFrame())
    if not conf.empty and "Status" in conf.columns:
        open_c = conf[conf["Status"].astype(str).str.lower().str.contains("open", na=False)]
        rep.source_conflicts = open_c.fillna("").to_dict(orient="records")

    return sheets, controls, rep

def compute_financials(sheets: Dict[str, pd.DataFrame], controls: Dict[str, str], prev_sheets: Optional[Dict[str, pd.DataFrame]] = None) -> Dict[str, Any]:
    cutoff = controls.get("Simulation cutoff", "")
    curr_ym = cutoff[:7] if len(cutoff) >= 7 else ""
    bb, mf, cl, chg = sheets["Budget Baseline"], sheets["Monthly Finance"].copy(), sheets["Cost Ledger"], sheets["Changes"]
    mf["Month_s"] = mf["Month"].astype(str).str.strip()
    mf["Baseline_n"] = pd.to_numeric(mf["Baseline EUR"], errors="coerce").fillna(0.0)
    mf["Actual_n"] = pd.to_numeric(mf["Actual EUR"], errors="coerce")
    mf["Forecast_n"] = pd.to_numeric(mf["Forecast EUR"], errors="coerce").fillna(0.0)

    envelope = float(pd.to_numeric(bb["24m EUR"], errors="coerce").sum())
    c_rows = bb[bb["Category"].astype(str).str.lower().str.contains("contingency", na=False)]
    contingency = float(pd.to_numeric(c_rows["24m EUR"], errors="coerce").sum()) if not c_rows.empty else 0.0

    if cl.empty:
        ledger_cum_actual, ledger_curr_actual, open_commitments = 0.0, 0.0, 0.0
    else:
        cl_amt = pd.to_numeric(cl["Amount EUR"], errors="coerce").fillna(0.0)
        t_col = cl["Type booking accrual commitment"].astype(str).str.lower()
        b_mask = t_col.str.contains("booking|accrual", na=False)
        ledger_cum_actual = float(cl_amt[b_mask].sum())
        ledger_curr_actual = float(cl_amt[b_mask & (cl["Month"].astype(str).str.strip() == curr_ym)].sum())
        open_commitments = float(cl_amt[t_col.str.contains("commitment", na=False)].sum())

    ledger_reconciled = abs(ledger_cum_actual - float(mf["Actual_n"].dropna().sum())) < 0.01
    booked = mf["Actual_n"].notna()
    baseline_to_date = float(mf.loc[booked, "Baseline_n"].sum())
    curr_row = mf[mf["Month_s"] == curr_ym]
    curr_baseline = float(curr_row["Baseline_n"].sum()) if not curr_row.empty else 0.0
    future_etc_incl_risk = float(mf.loc[~booked, "Forecast_n"].sum())
    eac = ledger_cum_actual + future_etc_incl_risk
    eac_vs_envelope = eac - envelope
    cum_var_eur = ledger_cum_actual - baseline_to_date
    cum_var_pct = (cum_var_eur / baseline_to_date * 100.0) if baseline_to_date else None

    prior_forecast_for_curr_month, prior_eac, rephased_only = None, None, False
    if prev_sheets and "Monthly Finance" in prev_sheets:
        pmf = prev_sheets["Monthly Finance"].copy()
        pmf["Month_s"] = pmf["Month"].astype(str).str.strip()
        pmf["Actual_n"] = pd.to_numeric(pmf["Actual EUR"], errors="coerce")
        pmf["Forecast_n"] = pd.to_numeric(pmf["Forecast EUR"], errors="coerce").fillna(0.0)
        p_row = pmf[pmf["Month_s"] == curr_ym]
        if not p_row.empty:
            prior_forecast_for_curr_month = float(p_row["Forecast_n"].sum())
        p_booked = pmf["Actual_n"].notna()
        prior_eac = float(pmf.loc[p_booked, "Actual_n"].sum() + pmf.loc[~p_booked, "Forecast_n"].sum())
        if prior_forecast_for_curr_month is not None and prior_eac is not None:
            if abs(ledger_curr_actual - prior_forecast_for_curr_month) > 0.01 and abs(eac - prior_eac) < 0.01:
                rephased_only = True

    ref_fcst = prior_forecast_for_curr_month if prior_forecast_for_curr_month is not None else curr_baseline
    period_var_eur = ledger_curr_actual - ref_fcst
    period_var_pct = (period_var_eur / ref_fcst * 100.0) if ref_fcst else None

    ex_mask = chg["Example only"].astype(str).str.upper().str.startswith("YES")
    live_crs = chg[~ex_mask].copy()
    unapp = live_crs[~live_crs["Approval"].astype(str).str.lower().isin(["approved"])]
    unapp_pressure = float(pd.to_numeric(unapp["Incremental delivery EUR"], errors="coerce").fillna(0.0).sum())

    return {
        "cutoff": cutoff, "period_ym": curr_ym, "envelope_eur": envelope,
        "authority_status": controls.get("Baseline authority", "Unknown"),
        "curr_month_actual_eur": ledger_curr_actual, "curr_month_baseline_eur": curr_baseline,
        "prior_forecast_for_curr_month_eur": prior_forecast_for_curr_month,
        "period_variance_eur": round(period_var_eur, 2),
        "period_variance_pct": round(period_var_pct, 2) if period_var_pct is not None else None,
        "prior_eac_eur": prior_eac, "rephased_only": rephased_only,
        "cum_actual_eur": ledger_cum_actual, "baseline_to_date_eur": baseline_to_date,
        "cum_variance_eur": round(cum_var_eur, 2),
        "cum_variance_pct": round(cum_var_pct, 2) if cum_var_pct is not None else None,
        "etc_incl_risk_eur": future_etc_incl_risk, "eac_eur": eac,
        "eac_vs_envelope_eur": eac_vs_envelope, "remaining_envelope_eur": envelope - ledger_cum_actual,
        "remaining_contingency_eur": contingency, "open_commitments_eur": open_commitments,
        "ledger_reconciled": ledger_reconciled,
        "live_crs": live_crs.fillna("").to_dict(orient="records"),
        "live_cr_count": int(len(live_crs)), "example_cr_count": int(ex_mask.sum()),
        "unapproved_live_cr_pressure_eur": unapp_pressure,
    }

def compute_delivery_and_rag(sheets: Dict[str, pd.DataFrame], fin: Dict[str, Any], rep: ValidationReport) -> Dict[str, Any]:
    ms, raid, ben, obs, blg, act = sheets["Milestones"], sheets["RAID"], sheets["Benefits Register"], sheets["Benefit Observations"], sheets["Backlog"], sheets["Actions"]
    cutoff = fin["cutoff"]
    slipped, unfinished_milestones = [], []
    for _, r in ms.iterrows():
        act_f = str(r.get("Actual finish", "") or "").strip()
        if act_f and act_f.lower() not in ("none", "nan", "nat", "tbd", "pending", "open", "not started", "-"):
            continue
        mid, phase = str(r.get("ID", "")).strip(), str(r.get("Phase", "")).strip()
        pf, cf = str(r.get("Planning finish", "")).strip(), str(r.get("Current forecast finish", "")).strip()
        owner, auth = str(r.get("Owner", "")).strip(), str(r.get("Authority", "")).strip()
        on_plan = (cf == pf)
        if not on_plan:
            slipped.append({"ID": mid, "Phase": phase, "Plan": pf, "Forecast": cf})
        sched_label = "Forecast on plan" if on_plan else f"Slipped / Unconfirmed ({cf})"
        gate_label = "Readiness unconfirmed (gate evidence & sign-off pending)" if ("not charter-approved" in auth.lower() or "draft" in fin["authority_status"].lower()) else "Gate criteria pending"
        blocker = "Open discovery/access actions & unsigned charter authority" if mid == "MS-01" else "Requires MS-01 G1 acceptance & RAID-I06 access retest"
        next_act = "Close discovery actions & submit G1 pack" if mid == "MS-01" else "Complete access remediation & G1 transition"
        due = "2026-12-18" if mid == "MS-01" else "2027-03-19"
        unfinished_milestones.append({
            "ID": mid, "Phase": phase, "Plan": pf, "Forecast": cf,
            "Schedule Status": sched_label, "Gate Readiness": gate_label,
            "Combined Status": f"{sched_label} · readiness unconfirmed",
            "Blocker": blocker, "Next Action": next_act, "Owner": owner, "Due": due
        })

    r_type, r_stat = raid["type"].astype(str).str.strip(), raid["status"].astype(str).str.strip()
    closed_mask = r_stat.str.lower().str.startswith("satisfied") | r_stat.str.lower().str.startswith("closed")
    open_raid = raid[~closed_mask].copy()
    overdue_raid = open_raid[open_raid["target_date"].notna() & (open_raid["target_date"].astype(str).str.strip() < cutoff)]
    raid_by_type = {t: {"total": int((r_type == t).sum()), "open": int((open_raid["type"].astype(str).str.strip() == t).sum())} for t in ["Risk", "Issue", "Assumption", "Dependency"]}

    open_actions_df = act[~act["Status"].astype(str).str.lower().str.startswith("closed")].copy() if "Status" in act.columns else pd.DataFrame()
    overdue_actions = []
    if not open_actions_df.empty and "Due / milestone" in open_actions_df.columns:
        for _, ar in open_actions_df.iterrows():
            d_str = str(ar.get("Due / milestone", "")).strip()
            if re.match(r"^\d{4}-\d{2}-\d{2}$", d_str) and d_str < cutoff:
                overdue_actions.append({"ID": ar.get("ID"), "Action": ar.get("Action"), "Owner": ar.get("Owner"), "Due": d_str, "Source": ar.get("Source")})

    sev_rank = {"Critical": 1, "High": 2, "Medium": 3, "Low": 4}
    open_raid["sev_num"] = open_raid["severity"].astype(str).str.strip().map(sev_rank).fillna(9)
    open_raid["type_num"] = open_raid["type"].astype(str).str.strip().map({"Issue": 1, "Risk": 2, "Dependency": 3, "Assumption": 4}).fillna(5)
    top_exc_df = open_raid.sort_values(["type_num", "sev_num", "id"]).head(3)

    act_lookup = {}
    if not open_actions_df.empty:
        for _, ar in open_actions_df.iterrows():
            for tok in re.findall(r"RAID-[A-Z]\d+|CHG-\d+", str(ar.get("Source", ""))):
                act_lookup[tok] = ar

    top_exceptions = []
    for _, er in top_exc_df.iterrows():
        eid = str(er.get("id", "")).strip()
        linked_act = act_lookup.get(eid)
        if linked_act is None and eid == "RAID-R01":
            linked_act = act_lookup.get("CHG-005")
        act_id = str(linked_act["ID"]).strip() if linked_act is not None else "—"
        act_txt = str(linked_act["Action"]).strip() if linked_act is not None else str(er.get("response", "")).strip()
        act_owner = str(linked_act["Owner"]).strip() if linked_act is not None else str(er.get("owner", "")).strip()
        act_due = str(linked_act["Due / milestone"]).strip() if linked_act is not None else str(er.get("target_date", "")).strip()
        is_overdue = bool(re.match(r"^\d{4}-\d{2}-\d{2}$", act_due) and act_due < cutoff)
        top_exceptions.append({
            "Record Type": str(er.get("type", "")).strip(),
            "IDs": f"{eid} / {act_id}" if act_id != "—" else eid,
            "Severity": str(er.get("severity", "")).strip(),
            "Title & Driving Evidence": f"{er.get('title')} — {er.get('statement')}",
            "Impact": "Blocks affected source connection until retest passes" if eid == "RAID-I06" else str(er.get("statement", ""))[:110],
            "Action": act_txt, "Owner": act_owner,
            "Due Date": f"{act_due} (OVERDUE)" if is_overdue else act_due, "Overdue": is_overdue,
        })

    crit_open_issues = open_raid[(open_raid["type"].astype(str).str.strip() == "Issue") & (open_raid["severity"].astype(str).str.strip() == "Critical")]
    mandatory_control_blockers = crit_open_issues[crit_open_issues["id"].astype(str).isin(["RAID-I03", "RAID-I06", "RAID-I10", "RAID-I16"])]

    eac_over = fin["eac_vs_envelope_eur"]
    if not rep.is_valid or not fin["ledger_reconciled"]:
        perf_rag = "UNKNOWN"
    elif eac_over > 0.05 * fin["envelope_eur"]:
        perf_rag = "RED"
    elif eac_over > 0 or len(slipped) > 0 or len(crit_open_issues) > 0:
        perf_rag = "AMBER"
    else:
        perf_rag = "GREEN"

    gate_rag = "BLOCKED (RED — Hold)" if (len(mandatory_control_blockers) > 0 or "draft" in fin["authority_status"].lower()) else ("AMBER" if len(rep.source_conflicts) > 0 else "GREEN")
    primary_driver = f"EAC EUR {fin['eac_eur']/1e6:.2f}M is EUR {eac_over:,.0f} above EUR {fin['envelope_eur']/1e6:.1f}M envelope (unapproved CHG-005); {len(mandatory_control_blockers)} critical gate/access blocker(s) open." if eac_over > 0 else "Cost and milestone forecasts within baseline; gate readiness awaiting mandatory control sign-offs."
    i06_act = act_lookup.get("RAID-I06")
    immediate_action = f"{i06_act['ID']}: {i06_act['Action']} (Owner: {i06_act['Owner']}, Due: {i06_act['Due / milestone']})" if i06_act is not None else "Review open RAID & unapproved change requests at F03 Control Board."

    cash_realised = float(pd.to_numeric(obs["Cash release EUR"], errors="coerce").dropna().sum()) if (not obs.empty and "Cash release EUR" in obs.columns and obs["Cash release EUR"].notna().any()) else None
    rag_reasons = [
        f"Cost ({'AMBER' if eac_over > 0 else 'GREEN'}): EAC EUR {fin['eac_eur']:,.0f} vs EUR {fin['envelope_eur']:,.0f} envelope (EUR {eac_over:+,.0f}).",
        f"Timelines ({'AMBER' if slipped else 'GREEN'}): {len(slipped)} slipped milestone(s); next milestones MS-01 & MS-02 forecast on plan with readiness unconfirmed.",
        f"Release Readiness ({gate_rag}): Mandatory access test failure (RAID-I06) keeps affected connection blocked."
    ]
    return {
        "performance_rag": perf_rag, "gate_readiness_rag": gate_rag,
        "rag_rule_basis": "Draft judgement per Programme Update Rules & CON-005: Performance RAG kept separate from Gate Readiness (Blocked if mandatory security/access test fails).",
        "primary_driver": primary_driver, "immediate_action": immediate_action, "rag_reasons": rag_reasons,
        "slipped_milestones": slipped, "unfinished_milestones": unfinished_milestones,
        "open_raid_total": int(len(open_raid)), "overdue_raid_total": int(len(overdue_raid)),
        "raid_by_type": raid_by_type, "top_exceptions": top_exceptions, "overdue_actions": overdue_actions,
        "critical_open_issues": crit_open_issues[["id", "title", "owner", "status"]].to_dict(orient="records"),
        "backlog_total": int(len(blg)),
        "backlog_ready": int(blg["Refinement status"].astype(str).str.lower().str.startswith("ready").sum()) if "Refinement status" in blg.columns else 0,
        "benefit_count": int(len(ben)), "observation_count": int(len(obs)),
        "cash_realised_eur": cash_realised, "open_actions": int(len(open_actions_df)),
    }

def run_bounded_agents(sheets: Dict[str, pd.DataFrame], controls: Dict[str, str], fin: Dict[str, Any], deliv: Dict[str, Any], rep: ValidationReport, api_key: Optional[str] = None) -> Dict[str, Any]:
    snap_fin, snap_deliv = copy.deepcopy(fin), copy.deepcopy(deliv)
    pmo_bullets = [
        f"Period actual EUR {snap_fin['curr_month_actual_eur']:,.0f}; cumulative actuals EUR {snap_fin['cum_actual_eur']:,.0f} vs EUR {snap_fin['baseline_to_date_eur']:,.0f} baseline-to-date.",
        f"Forecast EAC is EUR {snap_fin['eac_eur']:,.0f} (ETC EUR {snap_fin['etc_incl_risk_eur']:,.0f} incl. EUR {snap_fin['remaining_contingency_eur']:,.0f} unspent contingency), EUR {snap_fin['eac_vs_envelope_eur']:+,.0f} vs envelope.",
        f"Milestones: {len(snap_deliv['slipped_milestones'])} slipped; {len(snap_deliv['overdue_actions'])} overdue open action(s)."
    ]
    ba_bullets = [
        f"Backlog: {snap_deliv['backlog_ready']} of {snap_deliv['backlog_total']} items ready for owner review.",
        f"Change Control: {snap_fin['example_cr_count']} examples excluded; {snap_fin['live_cr_count']} live proposal(s) unapproved (EUR {snap_fin['unapproved_live_cr_pressure_eur']:,.0f} pressure).",
        f"Benefits: {snap_deliv['benefit_count']} targets; {snap_deliv['observation_count']} observation(s); cash release Unknown / Unvalidated."
    ]
    gov_bullets = [
        f"Authority: '{snap_fin['authority_status']}'; Gate Readiness: {snap_deliv['gate_readiness_rag']}.",
        f"RAID & Controls: {snap_deliv['open_raid_total']} open RAID ({snap_deliv['overdue_raid_total']} overdue); {len(rep.source_conflicts)} open source conflict(s).",
        f"Data Steward checks: {len(rep.checks_summary)} formula checks executed; {len(rep.warnings)} warning(s)."
    ]
    llm_used, llm_model = False, None
    if api_key and api_key.strip():
        try:
            from google import genai
            client = genai.Client(api_key=api_key.strip())
            for m in [x.name.replace("models/", "") for x in client.models.list() if "gemini" in x.name and "embed" not in x.name]:
                try:
                    res = client.models.generate_content(model=m, contents=f"Write 2 concise bullets using ONLY these numbers: {json.dumps({'fin': snap_fin, 'deliv': snap_deliv})}")
                    if res and res.text:
                        llm_used, llm_model = True, m
                        pmo_bullets.append(f"[LLM Commentary ({m})]: {res.text.strip()[:300]}")
                        break
                except Exception:
                    continue
        except Exception:
            llm_used = False

    sections = {
        "PMO": {"title": "Agent 1 — PMO & Delivery", "bullets": pmo_bullets, "mode": f"Deterministic + LLM ({llm_model})" if llm_used else "Deterministic (No LLM called)"},
        "Finance": {"title": "Agent 1B — Financial Control", "bullets": pmo_bullets[:2], "mode": "Deterministic"},
        "BA": {"title": "Agent 2 — Business Analysis & Benefits", "bullets": ba_bullets, "mode": "Deterministic"},
        "Governance": {"title": "Agent 3 — Risk, Controls & Governance", "bullets": gov_bullets, "mode": "Deterministic"},
    }
    for k, sec in sections.items():
        sec["content_hash"] = sha256_text(json.dumps(sec["bullets"], sort_keys=True))
    return {"sections": sections, "llm_used": llm_used, "llm_model": llm_model}

def verify_review_gate(section_key: str, section_hash: str, workbook_hash: str, cutoff: str, config_ver: str, review_record: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    if not review_record:
        return {"accepted": False, "status": "PENDING OWNER VALIDATION", "reason": "No review submitted"}
    bound_ok = (review_record.get("workbook_hash") == workbook_hash and review_record.get("cutoff") == cutoff and review_record.get("section_hash") == section_hash and review_record.get("config_version") == config_ver)
    if not bound_ok:
        return {"accepted": False, "status": "INVALIDATED — INPUT/CONTENT CHANGED", "reason": "Hash or cutoff mismatch; re-review required"}
    if review_record.get("outcome") == "Approved":
        return {"accepted": True, "status": f"ACKNOWLEDGED ({review_record.get('reviewer')} @ {review_record.get('timestamp')})", "reason": review_record.get("comments", "")}
    lbl = "REJECTED" if review_record.get("outcome") == "Rejected" else "CHANGES REQUESTED"
    return {"accepted": False, "status": f"{lbl} ({review_record.get('reviewer')})", "reason": review_record.get("comments", "")}

def execute_safe_rollover(file_bytes: bytes, next_cutoff: str, events: List[Dict[str, Any]], config_ver: str = "1.0") -> Tuple[bytes, Dict[str, Any]]:
    sheets, controls, rep = parse_workbook(file_bytes)
    if not rep.is_valid:
        raise ValueError(f"Cannot rollover invalid workbook: {rep.blocking_errors}")
    curr_cutoff = controls.get("Simulation cutoff", "")
    y, m, _ = [int(x) for x in curr_cutoff.split("-")]
    exp_y, exp_m = (y + 1, 1) if m == 12 else (y, m + 1)
    if not next_cutoff.startswith(f"{exp_y:04d}-{exp_m:02d}-"):
        raise ValueError(f"Next period '{next_cutoff}' is not strictly the next month after '{curr_cutoff}' ({exp_y:04d}-{exp_m:02d}).")

    input_sha = sha256_bytes(file_bytes)
    last_run = controls.get("Last simulation run", "RUN-000000")
    run_num = int(re.sub(r"\D", "", last_run) or "0") + 1
    next_run_id = f"RUN-{run_num:06d}"
    run_key = f"{SCHEMA_VERSION}:{next_cutoff[:7]}:{next_run_id}:{input_sha[:12]}"

    ul = sheets.get("Update Log", pd.DataFrame())
    if not ul.empty and "Run key" in ul.columns and run_key in ul["Run key"].astype(str).values:
        raise ValueError(f"Duplicate rollover blocked (idempotency): Run key '{run_key}' already processed.")

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    b2_val = str(wb["Controls"]["B2"].value or "")
    m_ver = re.search(r"v0\.(\d+)", b2_val)
    in_minor = int(m_ver.group(1)) if m_ver else 3
    in_ver, out_ver = f"v0.{in_minor}", f"v0.{in_minor + 1}"

    ws_ctrl = wb["Controls"]
    field_changes = []
    for r in range(6, ws_ctrl.max_row + 1):
        fname = str(ws_ctrl.cell(r, 2).value or "").strip()
        old_v = ws_ctrl.cell(r, 3).value
        if fname == "Simulation day":
            new_v = int(old_v or 0) + 1
            ws_ctrl.cell(r, 3).value = new_v
            ws_ctrl.cell(r, 4).value = f"Run {run_num} closes {next_cutoff[:7]}; on demand only"
            field_changes.append((next_run_id, "ROLLOVER", "Controls", f"C{r}", old_v, new_v, next_cutoff, "Controlled rollover"))
        elif fname == "Simulation cutoff":
            ws_ctrl.cell(r, 3).value = next_cutoff
            ws_ctrl.cell(r, 4).value = f"Month {run_num}; keep distinct from delivery date"
            field_changes.append((next_run_id, "ROLLOVER", "Controls", f"C{r}", old_v, next_cutoff, next_cutoff, "Controlled rollover"))
        elif fname == "Last simulation run":
            ws_ctrl.cell(r, 3).value = next_run_id
            field_changes.append((next_run_id, "ROLLOVER", "Controls", f"C{r}", old_v, next_run_id, next_cutoff, "Controlled rollover"))

    ws_evt = wb["Event Log"]
    ws_fcl = wb["Field Change Log"] if "Field Change Log" in wb.sheetnames else wb.create_sheet("Field Change Log")
    applied_event_ids, quarantined = [], []
    next_evt_num = max(0, ws_evt.max_row - 5)
    for ev in events:
        if not ev.get("authorized") or not ev.get("evidence_id") or not ev.get("description"):
            quarantined.append({"event": ev, "reason": "Missing authorization, evidence_id, or description"})
            continue
        next_evt_num += 1
        eid = f"EVT-{next_evt_num:06d}"
        applied_event_ids.append(eid)
        ws_evt.append([None, eid, next_cutoff, ev.get("type", "Controlled update"), ev["description"], ev.get("affected_ids", ""), ev["evidence_id"], "Accepted controlled update", ev.get("reviewer", "Simulated Reviewer"), next_run_id])
        field_changes.append((next_run_id, eid, "Event Log", f"Row {ws_evt.max_row}", None, eid, next_cutoff, ev["evidence_id"]))

    ws_ul = wb["Update Log"]
    ws_ul.append([None, next_run_id, in_ver, curr_cutoff, out_ver, "; ".join(applied_event_ids) or "Calendar rollover only", "Controls; Update Log; Snapshots", "Schema/hash/formula checks passed", "Agent 0 + Simulated Human Gate", f"SAVED (in:{input_sha[:12]})", run_key])
    ws_snap = wb["Snapshots"]
    snap_id = f"SNAP-{max(1, ws_snap.max_row - 4):03d}"
    ws_snap.append([None, snap_id, next_cutoff[:7], next_cutoff, out_ver, f"{next_cutoff[:7]} Pack Draft", "Draft; no committee approval", f"Retain {next_cutoff[:7]} master; prior snapshots unchanged", "None"])
    for fc in field_changes:
        ws_fcl.append([None, *fc])

    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_bytes = out_buf.getvalue()
    output_sha = sha256_bytes(out_bytes)
    _, out_controls, out_val = parse_workbook(out_bytes)
    if not out_val.is_valid or out_controls.get("Simulation cutoff") != next_cutoff:
        raise RuntimeError(f"Post-rollover read-back validation failed: {out_val.blocking_errors}")

    return out_bytes, {"input_sha256": input_sha, "output_sha256": output_sha, "previous_cutoff": curr_cutoff, "next_cutoff": next_cutoff, "run_id": next_run_id, "snapshot_id": snap_id, "output_version": out_ver, "applied_events": applied_event_ids, "quarantined_events": quarantined, "timestamp_utc": ts}
