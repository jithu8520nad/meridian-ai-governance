import io, json, zipfile
import pandas as pd
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

NAVY, GOLD, SLATE, WHITE = RGBColor(15, 36, 56), RGBColor(197, 155, 39), RGBColor(44, 62, 80), RGBColor(255, 255, 255)
LIGHT_BG, RED_ALERT, GREEN_OK = RGBColor(244, 246, 249), RGBColor(176, 42, 55), RGBColor(25, 135, 84)

def _excel_bytes(sheets_dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    return buf.getvalue()

def build_cost_analysis_xlsx(sheets):
    mf = sheets["Monthly Finance"].copy()
    mf["Actual EUR"] = pd.to_numeric(mf["Actual EUR"], errors="coerce")
    mf["Forecast EUR"] = pd.to_numeric(mf["Forecast EUR"], errors="coerce")
    mf["Baseline EUR"] = pd.to_numeric(mf["Baseline EUR"], errors="coerce")
    mf["Actual_minus_Forecast"] = mf["Actual EUR"] - mf["Forecast EUR"]
    mf["Variance_Pct"] = (mf["Actual_minus_Forecast"] / mf["Forecast EUR"] * 100.0).round(2)
    vert = mf[["Month", "Baseline EUR", "Forecast EUR", "Actual EUR", "Actual_minus_Forecast", "Variance_Pct", "Evidence status"]].copy()
    totals = pd.DataFrame([{"Month": "TOTAL", "Baseline EUR": vert["Baseline EUR"].sum(), "Forecast EUR": vert["Forecast EUR"].sum(), "Actual EUR": vert["Actual EUR"].dropna().sum(), "Actual_minus_Forecast": vert["Actual_minus_Forecast"].dropna().sum(), "Variance_Pct": None, "Evidence status": "Reconciled to Cost Ledger"}])
    vert_out = pd.concat([vert, totals], ignore_index=True)
    horiz_cols = {"Metric": ["Baseline EUR", "Forecast EUR", "Actual EUR", "Actual - Forecast"]}
    for _, r in mf.iterrows():
        horiz_cols[str(r["Month"])] = [r["Baseline EUR"], r["Forecast EUR"], r["Actual EUR"], r["Actual_minus_Forecast"]]
    horiz = pd.DataFrame(horiz_cols)
    horiz["Total_EUR"] = horiz.iloc[:, 1:].sum(axis=1, numeric_only=True)
    return _excel_bytes({"Horizontal_Monthly_View": horiz, "Vertical_Comparison": vert_out, "Cost_Ledger": sheets["Cost Ledger"]})

def build_raid_xlsx(sheets):
    raid = sheets["RAID"].copy()
    r_type = raid["type"].astype(str).str.strip()
    return _excel_bytes({"Risks_Only": raid[r_type == "Risk"], "Issues_Only": raid[r_type == "Issue"], "Assumptions_Dependencies": raid[r_type.isin(["Assumption", "Dependency"])], "Open_Conflicts": sheets["Conflicts"]})

def build_backlog_xlsx(sheets):
    blg = sheets["Backlog"].copy()
    cols = [c for c in ["ID", "Epic", "Solution", "Item", "Refinement status", "Impact 1–5", "Effort 1–5", "Priority", "Mandatory control", "Delivery authority", "Owner"] if c in blg.columns]
    return _excel_bytes({"Backlog_Review": blg[cols] if cols else blg})

def build_benefits_xlsx(sheets):
    return _excel_bytes({"Benefits_Register": sheets["Benefits Register"], "Benefit_Observations": sheets["Benefit Observations"]})

def build_change_summary_docx(meta, sheets, review_gates):
    doc = Document()
    doc.add_heading("MERIDIAN RETAIL BANK — MONTHLY CHANGE SUMMARY (DRAFT)", level=0)
    doc.add_paragraph(f"Cutoff: {meta['cutoff']} | Run ID: {meta['run_id']} | SHA-256: {meta['wb_hash'][:16]} | Status: {meta['overall_review_label']}")
    for _, r in sheets["Changes"].iterrows():
        doc.add_paragraph(f"{r.get('ID')} [{r.get('Type')}] — {r.get('Title')} | Owner: {r.get('Owner')} | Delta EUR: {r.get('Incremental delivery EUR')} | Approval: {r.get('Approval')}", style="List Bullet")
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

def build_monthly_status_docx(meta, fin, deliv, agent_out, review_gates):
    doc = Document()
    doc.add_heading("MERIDIAN RETAIL BANK — MONTHLY STATUS REPORT (DRAFT)", level=0)
    doc.add_paragraph(f"Cutoff: {meta['cutoff']} | Run ID: {meta['run_id']} | Schema: {meta['schema']} | SHA-256: {meta['wb_hash'][:16]} | {meta['overall_review_label']}")
    doc.add_heading(f"1. RAG Summary — Performance: {deliv['performance_rag']} | Release Readiness: {deliv['gate_readiness_rag']}", level=1)
    for r in deliv["rag_reasons"]:
        doc.add_paragraph(r, style="List Bullet")
    doc.add_heading("2. Financial & Delivery KPIs", level=1)
    doc.add_paragraph(f"Envelope: EUR {fin['envelope_eur']:,.0f} ({fin['authority_status']}) | Period Actual: EUR {fin['curr_month_actual_eur']:,.0f} | Cum. Actuals: EUR {fin['cum_actual_eur']:,.0f} | EAC: EUR {fin['eac_eur']:,.0f}")
    doc.add_heading("3. Role Sections (Human Review Gated)", level=1)
    for k, sec in agent_out["sections"].items():
        g = review_gates[k]
        doc.add_heading(f"{sec['title']} — {g['status']}", level=2)
        for b in (sec["bullets"] if g["accepted"] else [f"WITHHELD / PENDING: {g['reason']}"]):
            doc.add_paragraph(b, style="List Bullet")
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

def build_steerco_pptx(meta, fin, deliv, agent_out, review_gates):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    s1 = prs.slides.add_slide(blank)
    bg = s1.background.fill; bg.solid(); bg.fore_color.rgb = NAVY
    tf = s1.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.6), Inches(3.2)).text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = f"MERIDIAN RETAIL BANK — {meta['overall_review_label'].upper()}"; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = GOLD
    p2 = tf.add_paragraph(); p2.text = "Executive Review & Steering Committee Pack"; p2.font.size = Pt(34); p2.font.bold = True; p2.font.color.rgb = WHITE; p2.space_before = Pt(10)
    p3 = tf.add_paragraph(); p3.text = f"Cutoff: {meta['cutoff']} | Run: {meta['run_id']} | Performance RAG: {deliv['performance_rag']} | Release Readiness: {deliv['gate_readiness_rag']}"; p3.font.size = Pt(15); p3.font.color.rgb = RGBColor(210, 220, 230); p3.space_before = Pt(16)

    s2 = prs.slides.add_slide(blank)
    hdr = s2.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(1.0))
    hdr.fill.solid(); hdr.fill.fore_color.rgb = NAVY; hdr.line.fill.background()
    ht = s2.shapes.add_textbox(Inches(0.6), Inches(0.25), Inches(12), Inches(0.5)).text_frame.paragraphs[0]
    ht.text = "1. Data-Driven Programme Scorecard & RAG Drivers"; ht.font.size = Pt(21); ht.font.bold = True; ht.font.color.rgb = WHITE
    cards = [("PERIOD ACTUAL", f"EUR {fin['curr_month_actual_eur']:,.0f}", f"Var vs Forecast: EUR {fin['period_variance_eur']:+,.0f}"), ("FORECAST EAC", f"EUR {fin['eac_eur']:,.0f}", f"Var vs Env: EUR {fin['eac_vs_envelope_eur']:+,.0f}"), ("OPEN RAID / SLIPPAGE", f"{deliv['open_raid_total']} Open", f"{len(deliv['slipped_milestones'])} Slipped Milestone(s)"), ("REALISED CASH", "UNKNOWN", f"{deliv['observation_count']} Observation(s) Logged")]
    for idx, (lbl, val, sub) in enumerate(cards):
        c = s2.shapes.add_shape(1, Inches(0.6 + idx * 3.1), Inches(1.3), Inches(2.85), Inches(1.7))
        c.fill.solid(); c.fill.fore_color.rgb = LIGHT_BG; c.line.color.rgb = GOLD
        ctf = c.text_frame; ctf.word_wrap = True
        cp1 = ctf.paragraphs[0]; cp1.text = lbl; cp1.font.size = Pt(11); cp1.font.bold = True; cp1.font.color.rgb = SLATE
        cp2 = ctf.add_paragraph(); cp2.text = val; cp2.font.size = Pt(22); cp2.font.bold = True; cp2.font.color.rgb = NAVY; cp2.space_before = Pt(6)
        cp3 = ctf.add_paragraph(); cp3.text = sub; cp3.font.size = Pt(11); cp3.font.color.rgb = SLATE; cp3.space_before = Pt(4)
    ab = s2.shapes.add_textbox(Inches(0.6), Inches(3.3), Inches(12.1), Inches(3.6)).text_frame; ab.word_wrap = True
    ap0 = ab.paragraphs[0]; ap0.text = "RAG Drivers & Gate Blockers:"; ap0.font.size = Pt(16); ap0.font.bold = True; ap0.font.color.rgb = NAVY
    for r in deliv["rag_reasons"]:
        ap = ab.add_paragraph(); ap.text = f"• {r}"; ap.font.size = Pt(13); ap.font.color.rgb = RED_ALERT; ap.space_before = Pt(6)

    s3 = prs.slides.add_slide(blank)
    hdr3 = s3.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(1.0))
    hdr3.fill.solid(); hdr3.fill.fore_color.rgb = NAVY; hdr3.line.fill.background()
    ht3 = s3.shapes.add_textbox(Inches(0.6), Inches(0.25), Inches(12), Inches(0.5)).text_frame.paragraphs[0]
    ht3.text = "2. Human-Reviewed Domain Sections (Segregation of Duties)"; ht3.font.size = Pt(21); ht3.font.bold = True; ht3.font.color.rgb = WHITE
    for idx, (k, sec) in enumerate(agent_out["sections"].items()):
        g = review_gates[k]
        box = s3.shapes.add_shape(1, Inches(0.45 + idx * 3.15), Inches(1.3), Inches(2.95), Inches(5.6))
        box.fill.solid(); box.fill.fore_color.rgb = LIGHT_BG; box.line.color.rgb = GREEN_OK if g["accepted"] else RED_ALERT
        btf = box.text_frame; btf.word_wrap = True
        bp0 = btf.paragraphs[0]; bp0.text = sec["title"]; bp0.font.size = Pt(13); bp0.font.bold = True; bp0.font.color.rgb = NAVY
        bp1 = btf.add_paragraph(); bp1.text = g["status"]; bp1.font.size = Pt(10); bp1.font.bold = True; bp1.font.color.rgb = GREEN_OK if g["accepted"] else RED_ALERT; bp1.space_before = Pt(4)
        for b in (sec["bullets"] if g["accepted"] else [f"WITHHELD: {g['reason']}"]):
            bp = btf.add_paragraph(); bp.text = f"• {b}"; bp.font.size = Pt(10.5); bp.font.color.rgb = SLATE if g["accepted"] else RED_ALERT; bp.space_before = Pt(6)
    out = io.BytesIO(); prs.save(out); return out.getvalue()

def build_full_reporting_pack_zip(meta, sheets, fin, deliv, agent_out, review_gates, rollover_bytes=None, rollover_name=None):
    ym = fin["period_ym"] or "Period"
    files = {
        f"01_Monthly_Status_Report_{ym}.docx": build_monthly_status_docx(meta, fin, deliv, agent_out, review_gates),
        f"02_SteerCo_Deck_{ym}.pptx": build_steerco_pptx(meta, fin, deliv, agent_out, review_gates),
        f"03_RAID_Register_{ym}.xlsx": build_raid_xlsx(sheets),
        f"04_Cost_Analysis_{ym}.xlsx": build_cost_analysis_xlsx(sheets),
        f"05_Benefits_Analysis_{ym}.xlsx": build_benefits_xlsx(sheets),
        f"06_Backlog_Review_{ym}.xlsx": build_backlog_xlsx(sheets),
        f"07_Change_Summary_{ym}.docx": build_change_summary_docx(meta, sheets, review_gates),
        f"08_Role_Outputs_{ym}.json": json.dumps({"meta": meta, "financials": fin, "delivery": deliv, "roles": agent_out, "review_gates": review_gates}, indent=2, default=str).encode("utf-8"),
    }
    if rollover_bytes and rollover_name:
        files[f"09_{rollover_name}"] = rollover_bytes
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, content in files.items():
            zf.writestr(fname, content)
    return zbuf.getvalue(), files
