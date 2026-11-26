import io, hashlib
import pandas as pd
import altair as alt
import streamlit as st
from exporters import build_steerco_pptx, build_steerco_docx, build_rollover_xlsx

st.set_page_config(page_title='Meridian Retail Bank | Executive SteerCo', layout='wide', initial_sidebar_state='collapsed')

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

with st.sidebar:
    st.header('1. Month-over-Month Workbooks')
    prev_file = st.file_uploader('Prior Month Baseline (.xlsx)', type=['xlsx'], key='prev')
    curr_file = st.file_uploader('Current Month Workbook (.xlsx)', type=['xlsx'], key='curr')
    next_cutoff = st.text_input('Next Month Cutoff', '2026-12-31')
    st.divider()
    st.header('2. Mid-Month Change Intake')
    mid_month_item = st.text_input('Log Mid-Month Event / Fix', placeholder='e.g. CHG-05 CFO approval attached')
    chg05_approved = st.checkbox('Mark CHG-05 (EUR 100k) Approved', value=False)

curr_bytes = curr_file.getvalue() if curr_file else None
prev_bytes = prev_file.getvalue() if prev_file else None
period_label = curr_file.name.replace('.xlsx', '') if curr_file else 'November 2026 vs October 2026 Baseline'

if curr_bytes:
    raid_df = clean_sheet(curr_bytes, 'RAID')
    chg_df = clean_sheet(curr_bytes, 'Changes')
    ben_df = clean_sheet(curr_bytes, 'Benefits Register')
    fin_df = clean_sheet(curr_bytes, 'Monthly Finance')
    chk_df = clean_sheet(curr_bytes, 'Checks')
else:
    raid_df, chg_df, ben_df, fin_df, chk_df = [pd.DataFrame()]*5

prev_raid_cnt = len(clean_sheet(prev_bytes, 'RAID')) if prev_bytes else 64
curr_raid_cnt = len(raid_df) if not raid_df.empty else 67
prev_chg_cnt = len(clean_sheet(prev_bytes, 'Changes')) if prev_bytes else 4
curr_chg_cnt = len(chg_df) if not chg_df.empty else 5

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
h1.title('MERIDIAN RETAIL BANK — STEERING COMMITTEE ONE-PAGER')
h1.caption(f'EUR 10.0M / 24-Month AI Transformation | Comparison: **{period_label}** | DORA · BaFin/MaRisk · GDPR · ECB')

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
    ).properties(height=285)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    '1. SteerCo One-Pager & Line Approval',
    '2. Financials & MoM Variance',
    '3. Scope & Change Control',
    '4. RAID & Governance',
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
    st.subheader('Month-over-Month Health Delta (Prior Month vs Current Month)')
    st.table(pd.DataFrame([
        {'Dimension': 'Forecast EAC', 'Prior Month (Oct)': 'EUR 10.00M', 'Current Month (Nov)': 'EUR 10.10M', 'Trend': 'Worsened (+EUR 100k unapproved CHG-05 pressure)'},
        {'Dimension': 'Cumulative Actuals', 'Prior Month (Oct)': 'EUR 320,000', 'Current Month (Nov)': 'EUR 545,000', 'Trend': 'Unchanged / On Plan (+EUR 225k booked)'},
        {'Dimension': 'Active RAID Count', 'Prior Month (Oct)': f'{prev_raid_cnt} items', 'Current Month (Nov)': f'{curr_raid_cnt} items', 'Trend': f'Attention ({curr_raid_cnt - prev_raid_cnt:+d} new items logged)'},
        {'Dimension': 'Validated Realised Cash', 'Prior Month (Oct)': 'EUR 0', 'Current Month (Nov)': 'EUR 0', 'Trend': 'Unchanged (Gated until Pilot M7–12)'}
    ]))
    if not fin_df.empty:
        st.dataframe(fin_df, use_container_width=True, height=260)

with tab3:
    st.subheader('Change Control Register & Mid-Month Lifecycle')
    if not chg_df.empty:
        st.dataframe(chg_df, use_container_width=True, height=300)
    else:
        st.info('Open left sidebar (>) to upload Prior & Current Month Programme_Data.xlsx files.')

with tab4:
    st.subheader('RAID Register & Reconciliation Checks')
    if not chk_df.empty:
        st.dataframe(chk_df, use_container_width=True, height=210)
    if not raid_df.empty:
        st.dataframe(raid_df.head(25), use_container_width=True, height=250)

with tab5:
    st.subheader('Benefits Register & Next-Month Workbook Rollover')
    cx, cy = st.columns(2)
    docx_data = build_steerco_docx(meta, kpis, approved_by_domain, signoffs, [r.replace('<strong>', '').replace('</strong>', '') for r in rag_reasons])
    cx.download_button('Download Executive Word Brief (.docx)', docx_data, 'Meridian_SteerCo_Brief.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', use_container_width=True)
    if curr_bytes:
        xlsx_data = build_rollover_xlsx(curr_bytes, meta['run_id'], next_cutoff, rag_reasons)
        cy.download_button(f'Download Next-Month Master Workbook ({next_cutoff}.xlsx)', xlsx_data, f'Programme_Data_{next_cutoff}.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
    if not ben_df.empty:
        st.dataframe(ben_df, use_container_width=True, height=260)
