import streamlit as st
from audit import run_audit
from export_excel import export_to_excel

st.set_page_config(page_title='MAWB Audit Analyzer V5', layout='wide')
st.title('MAWB Audit Analyzer V5')

with st.sidebar:
    billing_file = st.file_uploader('Billing Charges Excel (.xlsx)', type=['xlsx'])
    eta_file = st.file_uploader('Optional MAWB / ETA / Branch Mapping (.xlsx)', type=['xlsx'])
    mawb_text = st.text_area('Optional MAWB Filter', height=120)
    low_thr = st.number_input('Low margin threshold', 0.0, 1.0, 0.30, 0.01)
    high_thr = st.number_input('High margin threshold', 0.0, 1.0, 0.80, 0.01)

if billing_file is None:
    st.info('Upload a Billing Charges Excel file to begin.')
    st.stop()

try:
    result = run_audit(billing_file, eta_file, mawb_text, float(low_thr), float(high_thr))
    if result.eta_parse_note:
        st.info(result.eta_parse_note)
    if result.mawb_keep and not result.mawb_not_found_df.empty:
        with st.expander('MAWB Not Found'):
            st.dataframe(result.mawb_not_found_df, use_container_width=True, hide_index=True)

    st.header('PAGE 1: CFO SUMMARY')
    c1, c2 = st.columns(2)
    with c1:
        st.subheader('KPI')
        st.dataframe(result.kpi_vertical, use_container_width=True, hide_index=True)
    with c2:
        st.subheader('Negative KPI')
        st.dataframe(result.neg_summary, use_container_width=True, hide_index=True)

    st.subheader('Charge Code Summary')
    st.dataframe(result.display_chargecode_summary, use_container_width=True, hide_index=True)
    st.subheader('Vendor Summary')
    st.dataframe(result.display_vendor_summary, use_container_width=True, hide_index=True)

    names = ['Exceptions','MAWB Summary','Client Summary','Margin Outliers','Negative Profit','Zero Margin','Zero Profit','Cost=Sell=0','Sell=0','Cost=0','ChargeCode Profit<0','Raw Data']
    tabs = st.tabs(names)
    frames = [
        result.display_exceptions, result.display_summary, result.display_client_summary,
        result.display_margin_outliers, result.display_negative_profit, result.display_zero_margin,
        result.display_zero_profit, result.display_both_zero, result.display_sell_zero_only,
        result.display_cost_zero_only, result.display_chargecode_profit_lt0_mawb, result.df
    ]
    for tab, frame in zip(tabs, frames):
        with tab:
            st.dataframe(frame, use_container_width=True, hide_index=True)

    st.divider()
    st.download_button('Download V5 Excel Report', export_to_excel(result), 'MAWB_Audit_Report_V5.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
except Exception as exc:
    st.exception(exc)
