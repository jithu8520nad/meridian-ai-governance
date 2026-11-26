import io, hashlib
import pandas as pd
import altair as alt
import streamlit as st
from exporters import build_steerco_pptx, build_steerco_docx, build_rollover_xlsx

st.set_page_config(page_title='Meridian Retail Bank | Executive SteerCo', layout='wide', initial_sidebar_state='collapsed')

st.markdown('''<style>
.block-container {padding-top: 2.3rem; padding-bottom: 0.6rem; max-width: 1420px;}
h1 {font-size: 1.45rem !important; margin-bottom: 0.25rem !important;}
h4 {font-size: 0.98rem !important; margin-bottom: 0.2rem !important; color: #F8FAFC !important;}
div[data-testid='stMetric'] {background: #132238; border-left: 4px solid #C59B27; padding: 8px 12px; border-radius: 4px; color: #FFFFFF;}
.rag-amber {background: #FFF8E7; border-left: 6px solid #D97706; padding: 11px 15px; border-radius: 4px; margin-bottom: 12px; color: #1F2937;}
.rag-green {background: #ECFDF5; border-left: 6px solid #059669; padding: 11px 15px; border-radius: 4px; margin-bottom: 12px; color: #1F2937;}
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

default_ms = pd.DataFrame([
    {'ID': 'MS-01', 'Phase': 'Foundations', 'Milestone': 'Programme Charter & Governance Setup', 'Plan Finish': '2026-10-31', 'Forecast Finish': '2026-10-31', 'Status': 'Completed', 'Owner': 'Nora Fisch'},
    {'ID': 'MS-02', 'Phase': 'Foundations', 'Milestone': 'Core Banking Data Lake Ingestion', 'Plan Finish': '2026-11-30', 'Forecast Finish': '2026-11-30', 'Status': 'On Track', 'Owner': 'Sarah Jenkins'},
    {'ID': 'MS-03', 'Phase': 'Governance', 'Milestone': 'AI CoE Mandate & DORA Control Catalogue', 'Plan Finish': '2027-03-31', 'Forecast Finish': '2027-03-31', 'Status': 'On Track', 'Owner': 'Marcus Vance'},
    {'ID': 'MS-04', 'Phase': 'Pilot', 'Milestone': 'Model Validation & Stress-Testing Gate', 'Plan Finish': '2027-06-30', 'Forecast Finish': '2027-07-10', 'Status': 'At Risk (+10d)', 'Owner': 'Dr. Aris Thorne'}
])

default_chg = pd.DataFrame([
    {'ID': 'CHG-001', 'Title': 'Add Retail Mortgage FAQ Corpus', 'Owner': 'Elena Rossi', 'Why Raised': 'Baseline knowledge expansion example', 'Cost EUR': 0, 'Schedule': '0d', 'Approval': 'Example Only'},
    {'ID': 'CHG-002', 'Title': 'Enhanced DORA ICT Incident Logging', 'Owner': 'Marcus Vance', 'Why Raised': 'Regulatory telemetry alignment example', 'Cost EUR': 25000, 'Schedule': '+5d', 'Approval': 'Example Only'},
    {'ID': 'CHG-003', 'Title': 'Vector Index Partitioning by Business Unit', 'Owner': 'Sarah Jenkins', 'Why Raised': 'GDPR data-segregation isolation example', 'Cost EUR': 40000, 'Schedule': '+7d', 'Approval': 'Example Only'},
    {'ID': 'CHG-004', 'Title': 'Parallel Model Fallback Routing', 'Owner': 'Dr. Aris Thorne', 'Why Raised': 'MaRisk operational resilience example', 'Cost EUR': 35000, 'Schedule': '0d', 'Approval': 'Example Only'},
    {'ID': 'CHG-005', 'Title': 'Additional Model Validation & Red-Teaming', 'Owner': 'Nora Fisch', 'Why Raised': 'Pre-pilot stress testing prior to M7 gate', 'Cost EUR': 100000, 'Schedule': '+10d', 'Approval': 'Not approved (Held)'}
])

default_fin = pd.DataFrame([
    {'Category': 'Actuals Booked (Oct–Nov)', 'Amount EUR': 545000, 'Envelope Share': '5.4%', 'Status': 'Reconciled in Ledger'},
    {'Category': 'Remaining Delivery ETC', 'Amount EUR': 8055000, 'Envelope Share': '79.8%', 'Status': 'Forecasted (Oct 2026–Sep 2028)'},
    {'Category': 'Remaining Risk Reserve', 'Amount EUR': 1500000, 'Envelope Share': '14.9%', 'Status': 'Held inside EUR 10.0M Envelope'},
    {'Category': 'CHG-05 Unapproved Pressure', 'Amount EUR': 100000, 'Envelope Share': '+1.0%', 'Status': 'Excluded from Committed Baseline'}
])

default_raid = pd.DataFrame([
    {'ID': 'RSK-01', 'Type': 'Risk', 'Severity': 'High', 'Title': 'Unapproved CHG-05 lifts EAC to ~EUR 10.1M', 'Owner': 'Nora Fisch', 'Status': 'Open', 'Mitigation': 'Held outside baseline pending SteerCo funding decision.'},
    {'ID': 'RSK-04', 'Type': 'Risk', 'Severity': 'Medium', 'Title': 'Stale metadata note in Controls!D8 (Month 1 vs Day 2)', 'Owner': 'Data Steward', 'Status': 'Open', 'Mitigation': 'Flagged by Agent 0; blocked from rollover logic.'},
    {'ID': 'RSK-09', 'Type': 'Risk', 'Severity': 'High', 'Title': 'DORA third-party ICT concentration & exit plan', 'Owner': 'Marcus Vance', 'Status': 'Assessed', 'Mitigation': 'Mandatory evidence gate before M6 sign-off.'},
    {'ID': 'DEP-03', 'Type': 'Dependency', 'Severity': 'High', 'Title': 'Pilot M7–12 matched baseline evidence for BEN-01', 'Owner': 'Elena Rossi', 'Status': 'Open', 'Mitigation': 'EUR 2.4M/yr cash remains conditional until verified.'}
])

default_ben = pd.DataFrame([
    {'ID': 'BEN-01', 'Benefit': 'Net manual effort & cash conversion', 'Owner': 'Nora Fisch / Elena Rossi', 'Target': 'EUR 2.4M/yr (40% touch-time cut)', 'Evidence Gate': 'Pilot M7–12', 'Realised EUR': 0},
    {'ID': 'BEN-02', 'Benefit': 'External service cost reduction', 'Owner': 'Nora Fisch / Victor Chen', 'Target': 'EUR 2.1M/yr (30% avoidable)', 'Evidence Gate': 'Scale M13–18', 'Realised EUR': 0},
    {'ID': 'BEN-03', 'Benefit': 'Avoided growth spend', 'Owner': 'Elena Rossi / Cost Owners', 'Target': 'EUR 1.5M/yr (30% avoided)', 'Evidence Gate': 'Scale M13–18', 'Realised EUR': 0},
    {'ID': 'BEN-04..08', 'Benefit': 'Productivity (+30%), CSAT, Governance & CoE', 'Owner': 'Programme Leads', 'Target': 'Non-cash operational KPIs', 'Evidence Gate': 'M6 / Pilot M7–12', 'Realised EUR': 0}
])

with st.sidebar:
    st.header('1. Month-over-Month Workbooks')
    prev_file = st.file_uploader('Prior Month Baseline (.xlsx)', type=['xlsx'], key='prev')
    curr_file = st.file_uploader('Current Month Workbook (.xlsx)', type=['xlsx'], key='curr')
    next_cutoff = st.text_input('Next Month Cutoff', '2026-12-31')
    st.divider()
    st.header('2. Mid-Month Controls')
    chg05_approved = st.checkbox('Simulate CHG-05 Approved & Risks Mitigated (Green RAG)', value=False)

curr_bytes = curr_file.getvalue() if curr_file else None
prev_bytes = prev_file.getvalue() if prev_file else None

if curr_bytes:
    ms_df = clean_sheet(curr_bytes, 'Milestones')
    raid_df = clean_sheet(curr_bytes, 'RAID')
    chg_df = clean_sheet(curr_bytes, 'Changes')
    ben_df = clean_sheet(curr_bytes, 'Benefits Register')
    fin_df = clean_sheet(curr_bytes, 'Monthly Finance')
    chk_df = clean_sheet(curr_bytes, 'Checks')
else:
    ms_df, raid_df, chg_df, ben_df, fin_df, chk_df = default_ms.copy(), default_raid.copy(), default_chg.copy(), default_ben.copy(), default_fin.copy(), pd.DataFrame()

prev_raid_cnt = len(clean_sheet(prev_bytes, 'RAID')) if prev_bytes else 64
curr_raid_cnt = len(clean_sheet(curr_bytes, 'RAID')) if curr_bytes else 67

if not chg05_approved:
    rag_status, rag_cls = 'AMBER', 'rag-amber'
    rag_reasons = [
        '<strong>Cost (Amber):</strong> Cumulative actuals are on plan (EUR 545k), but unapproved live proposal CHG-05 (+EUR 100k) lifts forecast EAC to ~EUR 10.1M against the EUR 10.0M envelope.',
        '<strong>Timelines (Green / Watch):</strong> Foundations milestones MS-01 and MS-02 are on schedule; Pilot Gate MS-04 carries a +10 day validation buffer watch.',
        '<strong>Risks (Amber):</strong> Active RAID rose from ' + str(prev_raid_cnt) + ' to ' + str(curr_raid_cnt) + ' items, driven by unapproved scope pressure, DORA exit-plan evidence, and stale metadata (Controls!D8).'
    ]
else:
    rag_status, rag_cls = 'GREEN', 'rag-green'
    rag_reasons = [
        '<strong>Cost (Green):</strong> Cumulative actuals (EUR 545k) and forecast EAC (EUR 10.0M) are fully reconciled within the approved envelope including the EUR 1.5M risk reserve.',
        '<strong>Timelines (Green):</strong> Critical path milestones across Foundations (M1–M6) and Pilot readiness (M7–M12) are tracking on schedule with zero slippage.',
        '<strong>Risks (Green):</strong> All High-severity RAID items have active mitigation owners and DORA, BaFin/MaRisk, GDPR & ECB control evidence verified.'
    ]

h1, h2 = st.columns([3.3, 1.2])
h1.title('MERIDIAN RETAIL BANK — STEERING COMMITTEE & AI GOVERNANCE')

st.markdown(f"<div class='{rag_cls}'><strong>OVERALL PORTFOLIO RAG STATUS: {rag_status}</strong><br>{'<br>'.join(rag_reasons)}</div>", unsafe_allow_html=True)

m1, m2, m3, m4 = st.columns(4)
m1.metric('Cost & Forecast EAC', 'EUR 10.10M' if not chg05_approved else 'EUR 10.00M', 'Actuals EUR 545k | +EUR 100k Pressure' if not chg05_approved else 'Within EUR 10.0M Envelope', delta_color='inverse' if not chg05_approved else 'normal')
m2.metric('Timelines & Milestones', '92% On Track' if not chg05_approved else '100% On Track', '11 On Plan | 1 Buffer Watch (+10d)' if not chg05_approved else 'All Milestones On Plan', delta_color='off')
m3.metric('Active Risks & RAID', f'{curr_raid_cnt} Items', f'{curr_raid_cnt - prev_raid_cnt:+d} vs Prior Month', delta_color='inverse' if curr_raid_cnt > prev_raid_cnt else 'normal')
m4.metric('Realised Cash Benefits', 'EUR 0', 'EUR 6.0M/yr Conditional (Pilot M7–12)', delta_color='off')

def make_donut(df_chart):
    return alt.Chart(df_chart).mark_arc(innerRadius=50, outerRadius=95).encode(
        theta=alt.Theta(field='Value', type='quantitative'),
        color=alt.Color(field='Category', type='nominal', scale=alt.Scale(range=['#38BDF8', '#F59E0B', '#64748B', '#EF4444']), legend=alt.Legend(orient='bottom', columns=1, labelLimit=250, title=None)),
        tooltip=['Category', 'Label']
    ).properties(height=255)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    '1. SteerCo One-Pager & Line Approval',
    '2. Financials & MoM Variance',
    '3. Milestones & Change Control',
    '4. Risks & Governance (RAID)',
    '5. Benefits & Next-Month Rollover'
])

with tab1:
    p1, p2, p3 = st.columns(3)
    with p1:
        st.markdown('#### 1. Cost & Envelope (EUR 10.1M EAC)')
        st.altair_chart(make_donut(pd.DataFrame([{'Category': 'Actuals Booked (EUR 0.55M)', 'Value': 0.545, 'Label': 'EUR 545,000'}, {'Category': 'Remaining ETC (EUR 8.06M)', 'Value': 8.055, 'Label': 'EUR 8,055,000'}, {'Category': 'Risk Reserve (EUR 1.50M)', 'Value': 1.500, 'Label': 'EUR 1,500,000'}])), use_container_width=True)
        with st.expander('📋 Click for Cost & Ledger Details'):
            st.dataframe(fin_df, use_container_width=True, height=180)
    with p2:
        st.markdown('#### 2. Timelines & Milestones (12 Total)')
        st.altair_chart(make_donut(pd.DataFrame([{'Category': 'Completed (3)', 'Value': 3, 'Label': '3 Milestones Done'}, {'Category': 'On Track (8)', 'Value': 8, 'Label': '8 Milestones On Plan'}, {'Category': 'Buffer Watch / At Risk (1)', 'Value': 1, 'Label': 'MS-04 (+10d Watch)'}])), use_container_width=True)
        with st.expander('📋 Click for Milestone Schedule Details'):
            st.dataframe(ms_df, use_container_width=True, height=180)
    with p3:
        st.markdown(f'#### 3. Risks & RAID Profile ({curr_raid_cnt} Items)')
        st.altair_chart(make_donut(pd.DataFrame([{'Category': 'High Severity (11)', 'Value': 11, 'Label': '11 High-Severity Items'}, {'Category': 'Medium Severity (34)', 'Value': 34, 'Label': '34 Medium-Severity Items'}, {'Category': 'Low / Closed (22)', 'Value': 22, 'Label': '22 Low / Closed Items'}])), use_container_width=True)
        with st.expander('📋 Click for Key Risks & Mitigations'):
            st.dataframe(raid_df.head(10), use_container_width=True, height=180)

    st.divider()
    d_ms, d_cr, d_ben = st.columns(3)
    with d_ms:
        st.markdown('**Milestones Summary**')
        st.dataframe(ms_df[['ID', 'Milestone', 'Status', 'Owner']], use_container_width=True, height=165)
    with d_cr:
        st.markdown('**Change Requests (5 Total)**')
        st.dataframe(chg_df[['ID', 'Title', 'Cost EUR', 'Approval']], use_container_width=True, height=165)
    with d_ben:
        st.markdown('**Benefits Register (EUR 6.0M/yr)**')
        st.dataframe(ben_df[['ID', 'Benefit', 'Target', 'Realised EUR']], use_container_width=True, height=165)

    st.subheader('PMO & Owner Line-by-Line Output Verification')
    draft_lines = [
        ('Cost & Finance', 'Cumulative actuals EUR 545,000 (+EUR 225k MoM); EAC ~EUR 10.1M incl. EUR 1.5M risk reserve.'),
        ('Timelines & CRs', 'Foundations on track; CHG-05 (EUR 100k / +10d) held as Not Approved pending SteerCo decision.'),
        ('Risks & Benefits', '67 RAID items tracked; EUR 6.0M/yr benefits gated until Pilot M7–12 ledger observations.')
    ]
    approved_by_domain = {'Cost & Finance': [], 'Timelines & CRs': [], 'Risks & Benefits': []}
    for idx, (dom, stmt) in enumerate(draft_lines):
        ca, cb, cc = st.columns([0.9, 3.2, 1.3])
        ca.markdown(f'**{dom}**')
        cb.write(stmt)
        decision = cc.selectbox('Status', ['Approved', 'Needs Changes', 'Wrong - Reject'], key=f'ln_{idx}', label_visibility='collapsed')
        if decision == 'Approved':
            approved_by_domain[dom].append(f'- {stmt}')

    meta = {'schema': 'BANK-AI-PMO-v1', 'cutoff': '2026-11-30', 'run_id': 'RUN-000003'}
    kpis = [('Cost & EAC', 'EUR 10.10M', 'Actuals EUR 545k'), ('Timelines', '92% On Track', '11 On Plan | 1 Watch'), ('Active RAID', f'{curr_raid_cnt} Items', f'{curr_raid_cnt - prev_raid_cnt:+d} MoM'), ('Realised Cash', 'EUR 0', 'EUR 6.0M/yr Conditional')]
    signoffs = {k: len(v) > 0 for k, v in approved_by_domain.items()}
    pptx_data = build_steerco_pptx(meta, kpis, approved_by_domain, signoffs, [r.replace('<strong>', '').replace('</strong>', '') for r in rag_reasons])
    h2.download_button('Export Boardroom Deck (.pptx)', pptx_data, 'Meridian_SteerCo_Deck.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation', type='primary', use_container_width=True)

with tab2:
    st.subheader('Month-over-Month Health Delta & Cost Ledger')
    st.table(pd.DataFrame([
        {'Pillar': 'Cost (EAC)', 'Prior Month (Oct)': 'EUR 10.00M', 'Current Month (Nov)': 'EUR 10.10M', 'Trend': 'Amber (+EUR 100k unapproved CHG-05 pressure)'},
        {'Pillar': 'Timelines', 'Prior Month (Oct)': '100% On Plan', 'Current Month (Nov)': '92% On Plan', 'Trend': 'Green/Watch (MS-04 +10d validation buffer)'},
        {'Pillar': 'Risks (RAID)', 'Prior Month (Oct)': f'{prev_raid_cnt} items', 'Current Month (Nov)': f'{curr_raid_cnt} items', 'Trend': f'Amber ({curr_raid_cnt - prev_raid_cnt:+d} new items logged)'}
    ]))
    st.dataframe(fin_df, use_container_width=True, height=240)

with tab3:
    st.subheader('Milestones & Change Control Database')
    st.dataframe(ms_df, use_container_width=True, height=210)
    st.dataframe(chg_df, use_container_width=True, height=250)

with tab4:
    st.subheader('Risks & Governance Register (RAID)')
    if not chk_df.empty:
        st.dataframe(chk_df, use_container_width=True, height=190)
    st.dataframe(raid_df.head(30), use_container_width=True, height=280)

with tab5:
    st.subheader('Benefits Register & Next-Month Workbook Rollover')
    cx, cy = st.columns(2)
    docx_data = build_steerco_docx(meta, kpis, approved_by_domain, signoffs, [r.replace('<strong>', '').replace('</strong>', '') for r in rag_reasons])
    cx.download_button('Download Executive Word Brief (.docx)', docx_data, 'Meridian_SteerCo_Brief.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', use_container_width=True)
    if curr_bytes:
        xlsx_data = build_rollover_xlsx(curr_bytes, meta['run_id'], next_cutoff, rag_reasons)
        cy.download_button(f'Download Next-Month Master Workbook ({next_cutoff}.xlsx)', xlsx_data, f'Programme_Data_{next_cutoff}.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
    else:
        cy.download_button('Download Full Excel Database (.xlsx)', to_excel_bytes({'Cost_Ledger': fin_df, 'Milestones': ms_df, 'Change_Requests': chg_df, 'RAID': raid_df, 'Benefits': ben_df}), 'Meridian_Programme_Database.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
    st.dataframe(ben_df, use_container_width=True, height=260)
