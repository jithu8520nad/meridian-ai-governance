import io, hashlib
import pandas as pd
import altair as alt
import streamlit as st
from exporters import build_steerco_pptx, build_steerco_docx, build_rollover_xlsx

st.set_page_config(page_title='Meridian Retail Bank | AI Transformation & SteerCo', layout='wide', initial_sidebar_state='collapsed')

st.markdown('''<style>
.block-container {padding-top: 2.3rem; padding-bottom: 0.5rem; max-width: 1420px;}
h1 {font-size: 1.45rem !important; margin-bottom: 0.1rem !important;}
h4 {font-size: 0.98rem !important; margin-bottom: 0.2rem !important; color: #F8FAFC !important;}
div[data-testid='stMetric'] {background: #132238; border-left: 4px solid #C59B27; padding: 8px 12px; border-radius: 4px; color: #FFFFFF;}
.rag-amber {background: #FFF8E7; border-left: 6px solid #D97706; padding: 10px 14px; border-radius: 4px; margin-bottom: 10px; color: #1F2937;}
.rag-green {background: #ECFDF5; border-left: 6px solid #059669; padding: 10px 14px; border-radius: 4px; margin-bottom: 10px; color: #1F2937;}
</style>''', unsafe_allow_html=True)

def clean_sheet(file_bytes, sheet_name):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    if sheet_name not in xl.sheet_names:
        return pd.DataFrame()
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=4)
    df = df.loc[:, ~df.columns.astype(str).str.startswith('Unnamed:')]
    return df.dropna(how='all')

def to_excel_bytes(sheets_dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for name, d in sheets_dict.items():
            d.to_excel(writer, sheet_name=name[:31], index=False)
    return buf.getvalue()

default_chg = pd.DataFrame([
    {'ID': 'CHG-001', 'Type': 'Scope Example', 'Title': 'Add Retail Mortgage FAQ Corpus', 'Requester': 'Retail Ops', 'Owner': 'Elena Rossi', 'Why Raised': 'Illustrative baseline example for low-risk knowledge expansion', 'Cost Impact EUR': 0, 'Schedule Impact': '0 days', 'Status': 'Example Only', 'Approval': 'Excluded from Live Forecast'},
    {'ID': 'CHG-002', 'Type': 'Control Example', 'Title': 'Enhanced DORA ICT Incident Logging', 'Requester': 'IT Risk', 'Owner': 'Marcus Vance', 'Why Raised': 'Illustrative example for regulatory telemetry alignment', 'Cost Impact EUR': 25000, 'Schedule Impact': '+5 days', 'Status': 'Example Only', 'Approval': 'Excluded from Live Forecast'},
    {'ID': 'CHG-003', 'Type': 'Architecture Example', 'Title': 'Vector Index Partitioning by Business Unit', 'Requester': 'Data Eng', 'Owner': 'Sarah Jenkins', 'Why Raised': 'Illustrative example for GDPR data-segregation isolation', 'Cost Impact EUR': 40000, 'Schedule Impact': '+7 days', 'Status': 'Example Only', 'Approval': 'Excluded from Live Forecast'},
    {'ID': 'CHG-004', 'Type': 'Vendor Example', 'Title': 'Parallel Model Fallback Routing', 'Requester': 'AI CoE', 'Owner': 'Dr. Aris Thorne', 'Why Raised': 'Illustrative resilience example under MaRisk', 'Cost Impact EUR': 35000, 'Schedule Impact': '0 days', 'Status': 'Example Only', 'Approval': 'Excluded from Live Forecast'},
    {'ID': 'CHG-005', 'Type': 'Live Forecast Pressure', 'Title': 'Additional Model Validation & Red-Teaming Capacity', 'Requester': 'Model Risk', 'Owner': 'Nora Fisch / Dr. Aris Thorne', 'Why Raised': 'Expanded pre-pilot stress testing required prior to M7 gate', 'Cost Impact EUR': 100000, 'Schedule Impact': '+10 days buffer', 'Status': 'Proposed (Live)', 'Approval': 'Not approved'}
])

default_fin = pd.DataFrame([
    {'Month': '2026-10 (M1)', 'Baseline EUR': 320000, 'Actual EUR': 320000, 'Forecast EUR': 320000, 'Variance EUR': 0, 'Evidence Status': 'Booked & Reconciled'},
    {'Month': '2026-11 (M2)', 'Baseline EUR': 225000, 'Actual EUR': 225000, 'Forecast EUR': 225000, 'Variance EUR': 0, 'Evidence Status': 'Booked (Cumulative EUR 545k)'},
    {'Month': '2026-12 (M3)', 'Baseline EUR': 310000, 'Actual EUR': None, 'Forecast EUR': 310000, 'Variance EUR': 'Not supplied', 'Evidence Status': 'Future Period (Open ETC)'},
    {'Month': 'Risk Reserve', 'Baseline EUR': 1500000, 'Actual EUR': 0, 'Forecast EUR': 1500000, 'Variance EUR': 0, 'Evidence Status': 'Held inside EUR 10.0M Envelope'}
])

default_raid = pd.DataFrame([
    {'ID': 'RSK-01', 'Type': 'Risk', 'Title': 'Unapproved CHG-05 lifts EAC above EUR 10.0M envelope', 'Owner': 'Nora Fisch', 'Severity': 'High', 'Status': 'Open', 'Impact & Mitigation': 'EUR 100k pressure held outside committed baseline until Steering Committee funding decision.'},
    {'ID': 'RSK-04', 'Type': 'Risk', 'Title': 'Stale metadata note in Controls!D8 (Month 1 vs Day 2)', 'Owner': 'Data Steward', 'Severity': 'Medium', 'Status': 'Open', 'Impact & Mitigation': 'Flagged automatically by Agent 0; blocked from driving rollover period logic.'},
    {'ID': 'RSK-09', 'Type': 'Risk', 'Title': 'DORA third-party ICT concentration & exit plan validation', 'Owner': 'Marcus Vance', 'Severity': 'High', 'Status': 'Assessed', 'Impact & Mitigation': 'Mandatory evidence gate prior to M6 architecture sign-off.'},
    {'ID': 'DEP-03', 'Type': 'Dependency', 'Title': 'Pilot M7–12 matched baseline observations for BEN-01', 'Owner': 'Elena Rossi', 'Severity': 'High', 'Status': 'Open', 'Impact & Mitigation': 'EUR 2.4M/yr cash conversion remains conditional until Finance verifies ledger release.'}
])

default_ben = pd.DataFrame([
    {'ID': 'BEN-01', 'Benefit': 'Net manual effort & cash conversion', 'Owner': 'Nora Fisch / Elena Rossi', 'Nature': 'Mixed: capacity + conditional cash', 'Modelled Target': 'EUR 2.4M/yr (40% touch-time cut)', 'First Evidence': 'Pilot M7–12', 'Realised Cash EUR': 0, 'Guardrail': 'EUR 16M capacity excluded from cash totals'},
    {'ID': 'BEN-02', 'Benefit': 'External service cost reduction', 'Owner': 'Nora Fisch / Victor Chen', 'Nature': 'Conditional hard cash', 'Modelled Target': 'EUR 2.1M/yr (30% avoidable)', 'First Evidence': 'Scale M13–18', 'Realised Cash EUR': 0, 'Guardrail': 'Excludes BEN-01 internal labour pool'},
    {'ID': 'BEN-03', 'Benefit': 'Avoided growth spend', 'Owner': 'Elena Rossi / Cost Owners', 'Nature': 'Conditional cost avoidance', 'Modelled Target': 'EUR 1.5M/yr (30% avoided)', 'First Evidence': 'Scale M13–18', 'Realised Cash EUR': 0, 'Guardrail': 'Requires approved counterfactual'},
    {'ID': 'BEN-04..08', 'Benefit': 'Productivity, CSAT, Governance, CoE & Adoption', 'Owner': 'Programme Leads', 'Nature': 'Non-cash operational KPIs', 'Modelled Target': '+30% productivity; >=95% evidence', 'First Evidence': 'M6 / Pilot M7–12', 'Realised Cash EUR': 0, 'Guardrail': 'Tracked separately; never double-counted'}
])

with st.sidebar:
    st.header('1. Month-over-Month Workbooks')
    prev_file = st.file_uploader('Prior Month Baseline (.xlsx)', type=['xlsx'], key='prev')
    curr_file = st.file_uploader('Current Month Workbook (.xlsx)', type=['xlsx'], key='curr')
    next_cutoff = st.text_input('Next Month Cutoff', '2026-12-31')
    st.divider()
    st.header('2. Mid-Month Change Intake')
    mid_month_item = st.text_input('Log Mid-Month Event / Fix', placeholder='e.g. CHG-05 CFO approval attached')
    chg05_approved = st.checkbox('Mark CHG-05 (EUR 100k) Approved', value=False)
    st.divider()
    st.link_button('Open Full Narrative Case Study', 'https://chatgpt.site', use_container_width=True)

curr_bytes = curr_file.getvalue() if curr_file else None
prev_bytes = prev_file.getvalue() if prev_file else None
period_label = curr_file.name.replace('.xlsx', '') if curr_file else 'November 2026 vs October 2026 Baseline'

if curr_bytes:
    raid_df = clean_sheet(curr_bytes, 'RAID')
    chg_df = clean_sheet(curr_bytes, 'Changes')
    chgd_df = clean_sheet(curr_bytes, 'Change Details')
    if not chg_df.empty and not chgd_df.empty and 'ID' in chg_df.columns and 'id' in chgd_df.columns:
        extra_cols = [c for c in ['id', 'problem', 'scope', 'timeline', 'benefits', 'recommendation', 'variance'] if c in chgd_df.columns]
        chg_df = chg_df.merge(chgd_df[extra_cols], left_on='ID', right_on='id', how='left').drop(columns=['id'], errors='ignore')
    ben_df = clean_sheet(curr_bytes, 'Benefits Register')
    fin_df = clean_sheet(curr_bytes, 'Monthly Finance')
    chk_df = clean_sheet(curr_bytes, 'Checks')
else:
    raid_df, chg_df, ben_df, fin_df, chk_df = default_raid.copy(), default_chg.copy(), default_ben.copy(), default_fin.copy(), pd.DataFrame()

prev_raid_cnt = len(clean_sheet(prev_bytes, 'RAID')) if prev_bytes else 64
curr_raid_cnt = len(clean_sheet(curr_bytes, 'RAID')) if curr_bytes else 67
curr_chg_cnt = len(chg_df)

if not chg05_approved:
    rag_status, rag_cls = 'AMBER', 'rag-amber'
    rag_reasons = [
        '1. <strong>Unapproved Forecast Pressure (CHG-05):</strong> Live EUR 100k proposal remains unapproved, lifting EAC to ~EUR 10.1M (+EUR 100k vs Prior Month).',
        '2. <strong>Conditional Benefits Only:</strong> EUR 6.0M/yr target benefits await Pilot M7–12 ledger evidence (EUR 0 realised cash — unchanged vs Prior Month).',
        '3. <strong>RAID & Pending Approvals:</strong> Active RAID moved from ' + str(prev_raid_cnt) + ' to ' + str(curr_raid_cnt) + ' items; line-item owner sign-off required before cutoff.'
    ]
else:
    rag_status, rag_cls = 'GREEN', 'rag-green'
    rag_reasons = ['1. <strong>Forecast Reconciled:</strong> CHG-05 funding decision logged under stage-gate authority.', '2. <strong>Controls Verified:</strong> DORA, BaFin/MaRisk, GDPR & ECB traceability intact.', '3. <strong>Ready for Cutoff:</strong> Approved line items locked for SteerCo release.']

if mid_month_item.strip():
    rag_reasons.append(f'4. <strong>Mid-Month Staged Update:</strong> {mid_month_item.strip()} (queued for month-end cutoff).')

h1, h2 = st.columns([3.3, 1.2])
h1.title('MERIDIAN RETAIL BANK — STEERING COMMITTEE & AI GOVERNANCE')
h1.caption('Jithu Nadviledath · MBA · PMP® · PRINCE2® Practitioner | EUR 10.0M / 24-Month Programme | DORA · BaFin/MaRisk · GDPR · ECB')

st.markdown(f"<div class='{rag_cls}'><strong>OVERALL PROGRAMME RAG STATUS: {rag_status}</strong><br>{'<br>'.join(rag_reasons)}</div>", unsafe_allow_html=True)

m1, m2, m3, m4 = st.columns(4)
m1.metric('Cumulative Actuals', 'EUR 545,000', '+EUR 225k vs Oct (On Plan)', delta_color='off')
m2.metric('Forecast EAC', 'EUR 10.10M' if not chg05_approved else 'EUR 10.00M', 'Worsened: +EUR 100k (CHG-05)' if not chg05_approved else 'Improved: Reconciled', delta_color='inverse' if not chg05_approved else 'normal')
m3.metric('Active RAID Items', f'{curr_raid_cnt} Items', f'{curr_raid_cnt - prev_raid_cnt:+d} vs Prior Month', delta_color='inverse' if curr_raid_cnt > prev_raid_cnt else 'normal')
m4.metric('Realised Cash Benefits', 'EUR 0', 'Unchanged (Pilot M7–12 Gated)', delta_color='off')

def make_donut(df_chart):
    return alt.Chart(df_chart).mark_arc(innerRadius=52, outerRadius=100).encode(
        theta=alt.Theta(field='Value', type='quantitative'),
        color=alt.Color(field='Category', type='nominal', scale=alt.Scale(range=['#38BDF8', '#F59E0B', '#64748B', '#EF4444']), legend=alt.Legend(orient='bottom', columns=1, labelLimit=250, title=None)),
        tooltip=['Category', 'Label']
    ).properties(height=275)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    '1. SteerCo One-Pager & Line Approval',
    '2. Financials & MoM Variance',
    '3. Scope & Change Control (5 CRs)',
    '4. RAID & Governance (67 Items)',
    '5. Benefits & Next-Month Rollover'
])

with tab1:
    p1, p2, p3 = st.columns(3)
    with p1:
        st.markdown('#### Cost & Forecast EAC (EUR 10.1M)')
        st.altair_chart(make_donut(pd.DataFrame([{'Category': 'Actuals (EUR 0.55M)', 'Value': 0.545, 'Label': 'EUR 545,000'}, {'Category': 'Remaining ETC (EUR 8.06M)', 'Value': 8.055, 'Label': 'EUR 8,055,000'}, {'Category': 'Risk Reserve (EUR 1.50M)', 'Value': 1.500, 'Label': 'EUR 1,500,000'}])), use_container_width=True)
    with p2:
        st.markdown(f'#### Change Requests ({curr_chg_cnt} Total)')
        st.altair_chart(make_donut(pd.DataFrame([{'Category': 'Baseline Examples (4)', 'Value': 4, 'Label': 'CHG-001 to CHG-004'}, {'Category': 'Unapproved Live CR (1)', 'Value': 1, 'Label': 'CHG-005 (EUR 100k Held)'}])), use_container_width=True)
    with p3:
        st.markdown('#### Conditional Benefits (EUR 6.0M/yr)')
        st.altair_chart(make_donut(pd.DataFrame([{'Category': 'BEN-01 Effort Cash (EUR 2.4M)', 'Value': 2.4, 'Label': 'EUR 2.4M/yr'}, {'Category': 'BEN-02 Service Cost (EUR 2.1M)', 'Value': 2.1, 'Label': 'EUR 2.1M/yr'}, {'Category': 'BEN-03 Avoided Spend (EUR 1.5M)', 'Value': 1.5, 'Label': 'EUR 1.5M/yr'}])), use_container_width=True)

    with st.expander('🔍 Quick Executive Drill-Down: View Change Requests, Costs & Benefits Database (or Download Excel)', expanded=False):
        db_bytes = to_excel_bytes({'Changes_Register': chg_df, 'Monthly_Finance': fin_df, 'RAID_Summary': raid_df, 'Benefits_Register': ben_df})
        st.download_button('📥 Download Complete Programme Database (.xlsx)', db_bytes, 'Meridian_Programme_Database.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        st.markdown('**Change Control Register (Owner, Rationale & Impact):**')
        st.dataframe(chg_df, use_container_width=True, height=215)

    st.subheader('PMO & Owner Line-by-Line Output Verification')
    draft_lines = [
        ('PMO & Finance', 'Cumulative actuals EUR 545,000 (+EUR 225k MoM); EAC ~EUR 10.1M incl. EUR 1.5M risk reserve.'),
        ('PMO & Finance', 'CHG-05 (EUR 100k forecast pressure) is Not Approved and excluded from committed baseline.'),
        ('Business Scope & Benefits', 'All 8 benefits (BEN-01..08, EUR 6.0M/yr) remain conditional until Pilot M7–12 observations are logged.'),
        ('Risk & Governance', 'Stale metadata note (Controls!D8) flagged; DORA, BaFin/MaRisk, GDPR & ECB audit trail preserved.')
    ]
    approved_by_domain = {'PMO & Finance': [], 'Business Scope & Benefits': [], 'Risk & Governance': []}
    for idx, (dom, stmt) in enumerate(draft_lines):
        ca, cb, cc = st.columns([0.9, 3.2, 1.3])
        ca.markdown(f'**{dom}**')
        cb.write(stmt)
        decision = cc.selectbox('Status', ['Approved', 'Needs Changes', 'Wrong - Reject'], key=f'ln_{idx}', label_visibility='collapsed')
        if decision == 'Approved':
            approved_by_domain[dom].append(f'- {stmt}')

    meta = {'schema': 'BANK-AI-PMO-v1', 'cutoff': '2026-11-30', 'run_id': 'RUN-000003'}
    kpis = [('Envelope', 'EUR 10.0M', '24-Month Baseline'), ('Cumulative Actuals', 'EUR 545k', '+EUR 225k MoM'), ('Active RAID', f'{curr_raid_cnt} Items', f'{curr_raid_cnt - prev_raid_cnt:+d} vs Prior Month'), ('Realised Cash', 'EUR 0', 'EUR 6.0M/yr Conditional')]
    signoffs = {k: len(v) > 0 for k, v in approved_by_domain.items()}
    pptx_data = build_steerco_pptx(meta, kpis, approved_by_domain, signoffs, [r.replace('<strong>', '').replace('</strong>', '') for r in rag_reasons])
    h2.download_button('Export Boardroom Deck (.pptx)', pptx_data, 'Meridian_SteerCo_Deck.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation', type='primary', use_container_width=True)

with tab2:
    st.subheader('Month-over-Month Health Delta & Financial Ledger')
    st.table(pd.DataFrame([
        {'Dimension': 'Forecast EAC', 'Prior Month (Oct)': 'EUR 10.00M', 'Current Month (Nov)': 'EUR 10.10M', 'Trend': 'Worsened (+EUR 100k unapproved CHG-05 pressure)'},
        {'Dimension': 'Cumulative Actuals', 'Prior Month (Oct)': 'EUR 320,000', 'Current Month (Nov)': 'EUR 545,000', 'Trend': 'Unchanged / On Plan (+EUR 225k booked)'},
        {'Dimension': 'Active RAID Count', 'Prior Month (Oct)': f'{prev_raid_cnt} items', 'Current Month (Nov)': f'{curr_raid_cnt} items', 'Trend': f'Attention ({curr_raid_cnt - prev_raid_cnt:+d} new items logged)'},
        {'Dimension': 'Validated Realised Cash', 'Prior Month (Oct)': 'EUR 0', 'Current Month (Nov)': 'EUR 0', 'Trend': 'Unchanged (Gated until Pilot M7–12)'}
    ]))
    st.dataframe(fin_df, use_container_width=True, height=250)

with tab3:
    st.subheader('Change Control Database (Owner, Rationale, Cost & Schedule Impact)')
    st.caption('Joined from Changes & Change Details sheets. CHG-001..004 are baseline examples; CHG-005 is the live EUR 100k proposal.')
    st.dataframe(chg_df, use_container_width=True, height=340)

with tab4:
    st.subheader('RAID Register & Reconciliation Checks')
    if not chk_df.empty:
        st.dataframe(chk_df, use_container_width=True, height=210)
    st.dataframe(raid_df.head(30), use_container_width=True, height=300)

with tab5:
    st.subheader('Benefits Register & Next-Month Workbook Rollover')
    cx, cy = st.columns(2)
    docx_data = build_steerco_docx(meta, kpis, approved_by_domain, signoffs, [r.replace('<strong>', '').replace('</strong>', '') for r in rag_reasons])
    cx.download_button('Download Executive Word Brief (.docx)', docx_data, 'Meridian_SteerCo_Brief.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', use_container_width=True)
    if curr_bytes:
        xlsx_data = build_rollover_xlsx(curr_bytes, meta['run_id'], next_cutoff, rag_reasons)
        cy.download_button(f'Download Next-Month Master Workbook ({next_cutoff}.xlsx)', xlsx_data, f'Programme_Data_{next_cutoff}.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
    else:
        cy.download_button('Download Baseline Excel Database (.xlsx)', to_excel_bytes({'Changes': chg_df, 'Finance': fin_df, 'RAID': raid_df, 'Benefits': ben_df}), 'Meridian_Baseline_Database.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
    st.dataframe(ben_df, use_container_width=True, height=280)
