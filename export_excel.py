from __future__ import annotations
import io
import pandas as pd
from helpers import to_date_only

PERCENT_FORMAT='0.00%'
NUMBER_FORMAT='#,##0.00'


def _safe_len(v):
    if v is None:
        return 0
    try:
        if pd.isna(v):
            return 0
    except Exception:
        pass
    return len(str(v))


def _format_sheet(ws, workbook, df):
    ws.freeze_panes(1,0)
    if len(df.columns):
        ws.autofilter(0,0,max(len(df),1),len(df.columns)-1)
    pf = workbook.add_format({'num_format':PERCENT_FORMAT})
    nf = workbook.add_format({'num_format':NUMBER_FORMAT})
    for i,col in enumerate(df.columns):
        width = min(max(max([len(str(col))]+[_safe_len(x) for x in df[col].head(200).tolist()])+2,11),42)
        if '%' in str(col):
            ws.set_column(i,i,width,pf)
        elif any(x in str(col) for x in ['Cost','Sell','Profit','Amount']):
            ws.set_column(i,i,width,nf)
        else:
            ws.set_column(i,i,width)


def export_to_excel(result) -> bytes:
    out = io.BytesIO()
    sheets = {
        'Raw_Data':to_date_only(result.df,['ETA']),
        'Exceptions':to_date_only(result.exceptions,['ETA']),
        'MAWB_Summary':to_date_only(result.summary,['ETA']),
        'Client_Summary':to_date_only(result.client_summary,['Latest_ETA']),
        'Margin_Outliers':to_date_only(result.margin_outliers,['ETA']),
        'Negative_Profit':to_date_only(result.negative_profit,['ETA']),
        'Zero_Margin':to_date_only(result.zero_margin,['ETA']),
        'Zero_Profit':to_date_only(result.zero_profit,['ETA']),
        'Cost=Sell=0':to_date_only(result.both_zero,['ETA']),
        'Sell=0':to_date_only(result.sell_zero_only,['ETA']),
        'Cost=0':to_date_only(result.cost_zero_only,['ETA']),
        'ChargeCode_Profit<0':to_date_only(result.chargecode_profit_lt0_mawb,['ETA']),
    }
    with pd.ExcelWriter(out,engine='xlsxwriter') as writer:
        wb = writer.book
        title = wb.add_format({'bold':True,'font_size':15})
        section = wb.add_format({'bold':True,'font_size':12})
        ws = wb.add_worksheet('Summary'); writer.sheets['Summary']=ws
        ws.write(0,0,'MAWB Audit Executive Summary',title)
        row=2
        for name,df in [('KPI',result.kpi_vertical),('Negative KPI',result.neg_summary),('Charge Code Summary',result.chargecode_summary),('Vendor Summary',result.vendor_summary)]:
            ws.write(row,0,name,section)
            df.to_excel(writer,sheet_name='Summary',index=False,startrow=row+1)
            row += len(df)+3
        ws.freeze_panes(2,0); ws.set_column(0,0,28); ws.set_column(1,20,18)
        for name,df in sheets.items():
            df.to_excel(writer,sheet_name=name,index=False)
            _format_sheet(writer.sheets[name],wb,df)
        if result.mawb_keep:
            result.mawb_not_found_df.to_excel(writer,sheet_name='MAWB_Not_Found',index=False)
            _format_sheet(writer.sheets['MAWB_Not_Found'],wb,result.mawb_not_found_df)
    out.seek(0)
    return out.getvalue()

build_excel_report = export_to_excel
