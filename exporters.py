import io, hashlib
from datetime import datetime, timezone
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from docx import Document
from docx.shared import Pt as DocPt, RGBColor as DocRGB

NAVY = RGBColor(15, 37, 55)
GOLD = RGBColor(197, 155, 39)
SLATE = RGBColor(44, 62, 80)
WHITE = RGBColor(255, 255, 255)
LIGHT_BG = RGBColor(244, 246, 249)
RED_ALERT = RGBColor(176, 42, 55)
GREEN_OK = RGBColor(25, 135, 84)

def build_steerco_pptx(meta, kpis, sections, signoffs, alerts):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    s1 = prs.slides.add_slide(blank)
    bg = s1.background.fill; bg.solid(); bg.fore_color.rgb = NAVY
    bar = s1.shapes.add_shape(1, Inches(0.8), Inches(1.4), Inches(1.2), Inches(0.08))
    bar.fill.solid(); bar.fill.fore_color.rgb = GOLD; bar.line.fill.background()
    tb = s1.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(11.5), Inches(2.2))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = 'MERIDIAN RETAIL BANK'; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = GOLD
    p2 = tf.add_paragraph(); p2.text = 'Enterprise AI Transformation — Steering Committee Pack'; p2.font.size = Pt(36); p2.font.bold = True; p2.font.color.rgb = WHITE; p2.space_before = Pt(12)
    p3 = tf.add_paragraph(); p3.text = f"Cutoff: {meta['cutoff']} | Schema: {meta['schema']} | Run: {meta['run_id']} | Governed by Banking SoD"; p3.font.size = Pt(15); p3.font.color.rgb = RGBColor(200, 210, 220); p3.space_before = Pt(18)

    s2 = prs.slides.add_slide(blank)
    hdr = s2.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(1.05))
    hdr.fill.solid(); hdr.fill.fore_color.rgb = NAVY; hdr.line.fill.background()
    ht = s2.shapes.add_textbox(Inches(0.6), Inches(0.28), Inches(12), Inches(0.5)).text_frame.paragraphs[0]
    ht.text = '1. Executive Programme & Financial Health Scorecard'; ht.font.size = Pt(22); ht.font.bold = True; ht.font.color.rgb = WHITE

    for idx, (label, val, sub) in enumerate(kpis[:4]):
        left = Inches(0.6 + idx * 3.1)
        card = s2.shapes.add_shape(1, left, Inches(1.4), Inches(2.85), Inches(1.75))
        card.fill.solid(); card.fill.fore_color.rgb = LIGHT_BG; card.line.color.rgb = GOLD
        ctf = card.text_frame; ctf.word_wrap = True
        cp1 = ctf.paragraphs[0]; cp1.text = label.upper(); cp1.font.size = Pt(11); cp1.font.bold = True; cp1.font.color.rgb = SLATE
        cp2 = ctf.add_paragraph(); cp2.text = str(val); cp2.font.size = Pt(24); cp2.font.bold = True; cp2.font.color.rgb = NAVY; cp2.space_before = Pt(8)
        cp3 = ctf.add_paragraph(); cp3.text = sub; cp3.font.size = Pt(11); cp3.font.color.rgb = SLATE; cp3.space_before = Pt(6)

    ab = s2.shapes.add_textbox(Inches(0.6), Inches(3.5), Inches(12.1), Inches(3.4)).text_frame; ab.word_wrap = True
    ap0 = ab.paragraphs[0]; ap0.text = 'Key Governance & Control Exceptions Flagged by Agent 0:'; ap0.font.size = Pt(16); ap0.font.bold = True; ap0.font.color.rgb = NAVY
    for a in alerts[:5]:
        ap = ab.add_paragraph(); ap.text = f'• {a}'; ap.font.size = Pt(13); ap.font.color.rgb = RED_ALERT; ap.space_before = Pt(6)

    s3 = prs.slides.add_slide(blank)
    hdr3 = s3.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(1.05))
    hdr3.fill.solid(); hdr3.fill.fore_color.rgb = NAVY; hdr3.line.fill.background()
    ht3 = s3.shapes.add_textbox(Inches(0.6), Inches(0.28), Inches(12), Inches(0.5)).text_frame.paragraphs[0]
    ht3.text = '2. Domain Briefings & Human-in-the-Loop Sign-Off Status'; ht3.font.size = Pt(22); ht3.font.bold = True; ht3.font.color.rgb = WHITE

    for idx, (role_name, bullets) in enumerate(sections.items()):
        left = Inches(0.6 + idx * 4.15)
        ok = signoffs.get(role_name, False)
        box = s3.shapes.add_shape(1, left, Inches(1.35), Inches(3.9), Inches(5.5))
        box.fill.solid(); box.fill.fore_color.rgb = LIGHT_BG; box.line.color.rgb = GREEN_OK if ok else RED_ALERT
        btf = box.text_frame; btf.word_wrap = True
        bp0 = btf.paragraphs[0]; bp0.text = role_name; bp0.font.size = Pt(16); bp0.font.bold = True; bp0.font.color.rgb = NAVY
        bp1 = btf.add_paragraph(); bp1.text = 'STATUS: VALIDATED BY OWNER' if ok else 'REDACTED: PENDING OWNER VALIDATION'; bp1.font.size = Pt(11); bp1.font.bold = True; bp1.font.color.rgb = GREEN_OK if ok else RED_ALERT; bp1.space_before = Pt(6)
        if ok:
            for b in bullets:
                bp = btf.add_paragraph(); bp.text = b.replace('**', ''); bp.font.size = Pt(12); bp.font.color.rgb = SLATE; bp.space_before = Pt(8)
        else:
            bp = btf.add_paragraph(); bp.text = 'Withheld from Steering Committee consolidation until designated human authority completes review and sign-off.'; bp.font.size = Pt(12); bp.font.color.rgb = RED_ALERT; bp.space_before = Pt(12)

    out = io.BytesIO(); prs.save(out); return out.getvalue()

def build_steerco_docx(meta, kpis, sections, signoffs, alerts):
    doc = Document()
    t = doc.add_heading('MERIDIAN RETAIL BANK — STEERING COMMITTEE BRIEF', level=0)
    t.runs[0].font.color.rgb = DocRGB(15, 37, 55)
    doc.add_paragraph(f"Programme: EUR 10M / 24-Month AI Transformation | Cutoff: {meta['cutoff']} | Schema: {meta['schema']} | Run ID: {meta['run_id']}")
    doc.add_heading('1. Executive Scorecard', level=1)
    tbl = doc.add_table(rows=1, cols=3); tbl.style = 'Table Grid'
    hdr = tbl.rows[0].cells; hdr[0].text = 'Metric'; hdr[1].text = 'Current Value'; hdr[2].text = 'Governance Basis'
    for label, val, sub in kpis:
        row = tbl.add_row().cells; row[0].text = label; row[1].text = str(val); row[2].text = sub
    doc.add_heading('2. Agent 0 Control & Quarantine Exceptions', level=1)
    for a in alerts:
        doc.add_paragraph(a, style='List Bullet')
    doc.add_heading('3. Domain Briefings (Human Sign-Off Gated)', level=1)
    for role_name, bullets in sections.items():
        ok = signoffs.get(role_name, False)
        doc.add_heading(f"{role_name} — {'VALIDATED BY OWNER' if ok else 'PENDING OWNER VALIDATION'}", level=2)
        if ok:
            for b in bullets:
                doc.add_paragraph(b.replace('**', '').lstrip('- '), style='List Bullet')
        else:
            doc.add_paragraph('REDACTED: Withheld pending designated authority validation.')
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

def build_rollover_xlsx(file_bytes, run_id, next_cutoff, alerts):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ts = datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')
    f_hash = hashlib.sha256(file_bytes).hexdigest()[:12]
    if 'Update Log' in wb.sheetnames:
        ws = wb['Update Log']
        ws.append([None, run_id, 'v0.3', next_cutoff, 'v0.4-ROLLED', 'CHG-05 (HELD)', 'BEN-01..08', 'PASS_WITH_FLAGS', 'Agent 0 + Human Gate', f'SAVED ({f_hash})', f'KEY-{run_id}'])
    if 'Field Change Log' in wb.sheetnames:
        ws2 = wb['Field Change Log']
        ws2.append([None, run_id, 'EVT-ROLLOVER', 'Controls', 'Simulation cutoff', '2026-11-30', next_cutoff, ts, f'Controlled rollover ({len(alerts)} flags preserved)'])
    out = io.BytesIO(); wb.save(out); return out.getvalue()
