from __future__ import annotations
from dataclasses import dataclass
import pandas as pd
import streamlit as st
from helpers import clean_eta_series, display_df, find_first_col, find_sheet_with_required_cols, format_pct_str, normalize_mawb, parse_mawb_list, pct, safe_numeric

BILLING_REQUIRED = {
    'MAWB': ['MAWB','Mawb','Master AWB','MasterAWB'],
    'Cost Amount': ['Cost Amount','Cost','AP Amount','Total Cost','CostAmount','AP'],
    'Sell Amount': ['Sell Amount','Sell','AR Amount','Total Sell','SellAmount','AR'],
}
BILLING_OPTIONAL = {
    'Client': ['Client','Customer','Account','Shipper','Bill To','Billed To'],
    'Charge Code': ['Charge Code','ChargeCode','Charge','Code'],
    'Vendor': ['Vendor','Carrier','Supplier'],
}
ETA_REQUIRED = {
    'MAWB': ['MAWB','Mawb','Master AWB','MasterAWB'],
    'ETA': ['ETA','Eta','Estimated Time of Arrival','Arrival','Arrival Date','ETA Date'],
}
ETA_OPTIONAL = {'Branch': ['Branch','Destination','Dest','Station']}
HANCAI_ALLOC_CODES = {'DTRF','TISC','TABD','DSTOR','WIO'}
HANCAI_THAWB_ONLY_CODES = {'THAWB','DDOC'}
HANCAI_SPECIAL_CLIENTS = {'HANCAIWUX','4PXDIGHKG'}
SHELIU_SPECIAL_CLIENTS = {'SHELIUSZX','LIBEXPLHR'}

@dataclass
class AuditResult:
    mawb_keep: list[str]
    mawb_not_found: list[str]
    mawb_not_found_df: pd.DataFrame
    eta_parse_note: str | None
    kpi_vertical: pd.DataFrame
    neg_summary: pd.DataFrame
    df: pd.DataFrame
    summary: pd.DataFrame
    exceptions: pd.DataFrame
    client_summary: pd.DataFrame
    margin_outliers: pd.DataFrame
    negative_profit: pd.DataFrame
    zero_margin: pd.DataFrame
    zero_profit: pd.DataFrame
    both_zero: pd.DataFrame
    sell_zero_only: pd.DataFrame
    cost_zero_only: pd.DataFrame
    chargecode_summary: pd.DataFrame
    vendor_summary: pd.DataFrame
    chargecode_profit_lt0_mawb: pd.DataFrame
    margin_label: str
    display_summary: pd.DataFrame
    display_exceptions: pd.DataFrame
    display_client_summary: pd.DataFrame
    display_margin_outliers: pd.DataFrame
    display_negative_profit: pd.DataFrame
    display_zero_margin: pd.DataFrame
    display_zero_profit: pd.DataFrame
    display_both_zero: pd.DataFrame
    display_sell_zero_only: pd.DataFrame
    display_cost_zero_only: pd.DataFrame
    display_chargecode_summary: pd.DataFrame
    display_vendor_summary: pd.DataFrame
    display_chargecode_profit_lt0_mawb: pd.DataFrame


def _clean_text(s: pd.Series, default='UNKNOWN') -> pd.Series:
    out = s.astype(str).str.strip().str.upper()
    return out.replace({'':default,'NAN':default,'NONE':default,'<NA>':default})


def _parts(mawb: str) -> tuple[str,str]:
    m = normalize_mawb(mawb)
    return (m[:3], m[-8:]) if len(m)==12 and m[3]=='-' else ('','')


def _build_procaresx_map(df: pd.DataFrame) -> dict[str,str]:
    pro = df.loc[df['Client'].eq('PROCARESX'), ['MAWB']].drop_duplicates()
    if pro.empty:
        return {}
    p = pro['MAWB'].apply(_parts)
    pro = pro.assign(Prefix=p.str[0], Last8=p.str[1])
    groups = pro[pro['Last8'].ne('')].groupby('Last8')['Prefix'].agg(set).to_dict()
    mapping = {}
    for last8, prefixes in groups.items():
        target = '125' if '125' in prefixes else ('932' if '932' in prefixes else ('001' if '001' in prefixes else ''))
        if target and '777' in prefixes:
            mapping[f'777-{last8}'] = f'{target}-{last8}'
    return mapping


def _apply_procaresx_merge(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str,str]]:
    out = df.copy()
    mapping = _build_procaresx_map(out)
    if mapping:
        mask = out['Client'].eq('PROCARESX')
        out.loc[mask, 'MAWB'] = out.loc[mask, 'MAWB'].replace(mapping)
    return out, mapping


def _apply_hancai_allocation(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out['Original Sell Amount'] = out['Sell Amount']
    out['AR Allocation Applied'] = False
    mask = out['Client'].eq('HANCAIWUX') & out['Charge Code'].isin(HANCAI_ALLOC_CODES)
    for _, idx in out.loc[mask].groupby('MAWB').groups.items():
        rows = out.loc[idx]
        if not rows['Charge Code'].eq('DTRF').any():
            continue
        total_ap = float(rows['Cost Amount'].sum())
        dtrf_ar = float(rows.loc[rows['Charge Code'].eq('DTRF'),'Sell Amount'].sum())
        if total_ap <= 0:
            continue
        out.loc[idx,'Sell Amount'] = dtrf_ar * rows['Cost Amount'] / total_ap
        out.loc[idx,'AR Allocation Applied'] = True
    return out


def _read_eta(eta_file):
    if eta_file is None:
        return None, None
    xls = pd.ExcelFile(eta_file)
    sh = find_sheet_with_required_cols(xls, ETA_REQUIRED)
    if not sh:
        return None, 'ETA mapping file uploaded, but no sheet with MAWB and ETA was found.'
    raw = pd.read_excel(xls, sheet_name=sh)
    mcol = find_first_col(raw, ETA_REQUIRED['MAWB'])
    ecol = find_first_col(raw, ETA_REQUIRED['ETA'])
    bcol = find_first_col(raw, ETA_OPTIONAL['Branch'])
    cols = [mcol, ecol] + ([bcol] if bcol else [])
    m = raw[cols].copy().rename(columns={mcol:'MAWB', ecol:'ETA', **({bcol:'Branch'} if bcol else {})})
    m['MAWB'] = m['MAWB'].apply(normalize_mawb)
    m['ETA'] = clean_eta_series(m['ETA'])
    m['Branch'] = m['Branch'].astype(str).str.strip().replace({'nan':'','None':''}) if 'Branch' in m.columns else ''
    bad = int(m['ETA'].isna().sum())
    note = f'ETA parsing note: {bad} / {len(m)} ETA values could not be parsed and were left blank.' if len(m) and bad else None
    m = m[m['MAWB'].ne('')].groupby('MAWB', as_index=False).agg(ETA=('ETA','max'), Branch=('Branch','last'))
    return m, note


def _active_codes(df: pd.DataFrame) -> dict[str,set[str]]:
    active = df.loc[df['Cost Amount'].ne(0) | df['Sell Amount'].ne(0), ['MAWB','Charge Code']].drop_duplicates()
    return active.groupby('MAWB')['Charge Code'].agg(set).to_dict()


def _classify(row, codes, low_thr, high_thr):
    client = str(row['Client']).upper()
    profit, pm = float(row['Profit']), float(row['Profit Margin %'])
    cost, sell = float(row['Total_Cost']), float(row['Total_Sell'])
    if profit < 0:
        return 'Exception','Open','Profit<0'
    if cost == 0 and sell == 0:
        return 'Exception','Open','Cost=Sell=0'
    if sell == 0:
        return 'Exception','Open','Revenue=0'
    if cost == 0:
        return 'Exception','Open','Cost=0'
    if client in HANCAI_SPECIAL_CLIENTS and codes and codes.issubset(HANCAI_THAWB_ONLY_CODES):
        return ('Exception','Open','Margin<85%') if pm < .85 else ('Exempt','Closed','THAWB/DDOC Margin>=85%')
    if client in SHELIU_SPECIAL_CLIENTS and 'THAWB' not in codes:
        return ('Exception','Open','Margin>35%') if pm > .35 else ('Exempt','Closed','No THAWB Margin<=35%')
    if pm > high_thr:
        return 'Exception','Open',f'Margin>{int(high_thr*100)}%'
    if pm < low_thr:
        return 'Exception','Open',f'Margin<{int(low_thr*100)}%'
    return 'Normal','Closed',''


def _pivot(relation, key, flags):
    joined = relation.drop_duplicates().merge(flags, on='MAWB', how='left')
    if joined.empty:
        return pd.DataFrame({key:[]})
    p = joined.pivot_table(index=key, columns='Exception_Type', values='MAWB', aggfunc=pd.Series.nunique, fill_value=0).reset_index()
    p.columns.name = None
    return p


def _kpi(metrics, pct_keys):
    return pd.DataFrame([{'Metric':k,'Value':format_pct_str(v) if k in pct_keys else v} for k,v in metrics.items()])


@st.cache_data(show_spinner=False)
def run_audit(billing_file, eta_file=None, mawb_text='', low_thr=.30, high_thr=.80) -> AuditResult:
    margin_label = f'Margin<{int(low_thr*100)}% or >{int(high_thr*100)}%'
    xls = pd.ExcelFile(billing_file)
    sh = find_sheet_with_required_cols(xls, BILLING_REQUIRED)
    if not sh:
        raise ValueError('Could not find a billing sheet containing MAWB, Cost/AP, and Sell/AR.')
    raw = pd.read_excel(xls, sheet_name=sh)
    mcol = find_first_col(raw, BILLING_REQUIRED['MAWB'])
    ccol = find_first_col(raw, BILLING_REQUIRED['Cost Amount'])
    scol = find_first_col(raw, BILLING_REQUIRED['Sell Amount'])
    clcol = find_first_col(raw, BILLING_OPTIONAL['Client'])
    chcol = find_first_col(raw, BILLING_OPTIONAL['Charge Code'])
    vcol = find_first_col(raw, BILLING_OPTIONAL['Vendor'])

    df = raw.copy()
    df['MAWB'] = df[mcol].apply(normalize_mawb)
    df['Original MAWB'] = df['MAWB']
    df['Cost Amount'] = safe_numeric(df[ccol])
    df['Sell Amount'] = safe_numeric(df[scol])
    df['Client'] = _clean_text(df[clcol]) if clcol else 'UNKNOWN'
    df['Charge Code'] = _clean_text(df[chcol]) if chcol else 'UNKNOWN'
    df['Vendor'] = _clean_text(df[vcol]) if vcol else 'UNKNOWN'
    df = df[df['MAWB'].ne('')].copy()

    df, pro_map = _apply_procaresx_merge(df)
    keep_raw = parse_mawb_list(mawb_text)
    mawb_keep = sorted({pro_map.get(x,x) for x in keep_raw})
    if mawb_keep:
        df = df[df['MAWB'].isin(mawb_keep)].copy()
        not_found = sorted(set(mawb_keep) - set(df['MAWB'].unique()))
    else:
        not_found = []
    not_found_df = pd.DataFrame({'MAWB':not_found})

    eta_map, eta_note = _read_eta(eta_file)
    if eta_map is not None and not eta_map.empty:
        eta_map['MAWB'] = eta_map['MAWB'].replace(pro_map)
        eta_map = eta_map.groupby('MAWB',as_index=False).agg(ETA=('ETA','max'),Branch=('Branch','last'))
        df = df.merge(eta_map,on='MAWB',how='left')
    else:
        df['ETA'] = pd.NaT
        df['Branch'] = ''
    df['ETA'] = pd.to_datetime(df['ETA'],errors='coerce').dt.normalize()
    df['Branch'] = df['Branch'].fillna('').astype(str).str.strip()

    df = _apply_hancai_allocation(df)
    df['Line Profit'] = df['Sell Amount'] - df['Cost Amount']
    code_map = _active_codes(df)

    summary = df.groupby('MAWB',as_index=False).agg(Client=('Client','first'),Branch=('Branch','first'),Total_Cost=('Cost Amount','sum'),Total_Sell=('Sell Amount','sum'),Line_Count=('MAWB','size'),ETA=('ETA','max'))
    summary['ETA Month'] = summary['ETA'].dt.to_period('M').astype(str).replace('NaT','')
    summary['Profit'] = summary['Total_Sell'] - summary['Total_Cost']
    summary['Profit Margin %'] = pct(summary['Profit'], summary['Total_Sell'])
    cls = summary.apply(lambda r:_classify(r,code_map.get(r['MAWB'],set()),low_thr,high_thr),axis=1,result_type='expand')
    cls.columns = ['Margin_Flag','Classification','Exception_Type']
    summary = pd.concat([summary,cls],axis=1)
    flags = summary[['MAWB','Margin_Flag','Classification','Exception_Type']]
    exceptions = summary[summary['Classification'].eq('Open')].copy().sort_values(['Exception_Type','Profit','MAWB'])

    client_summary = df.groupby('Client',as_index=False).agg(Total_Cost=('Cost Amount','sum'),Total_Sell=('Sell Amount','sum'),Line_Count=('Client','size'),MAWB_Count=('MAWB',pd.Series.nunique),Latest_ETA=('ETA','max'))
    client_summary['Profit'] = client_summary['Total_Sell'] - client_summary['Total_Cost']
    client_summary['Profit Margin %'] = pct(client_summary['Profit'],client_summary['Total_Sell'])
    client_summary = client_summary.sort_values('Profit',ascending=False)

    margin_outliers = summary[summary['Exception_Type'].str.startswith('Margin',na=False)].copy().sort_values('Profit Margin %')
    negative_profit = summary[summary['Profit'].lt(0)].copy().sort_values('Profit')
    zero_margin = summary[summary['Profit Margin %'].eq(0)].copy().sort_values(['Total_Sell','Total_Cost'],ascending=False)
    zero_profit = summary[summary['Profit'].eq(0)].copy().sort_values(['Total_Sell','Total_Cost'],ascending=False)
    both_zero = summary[summary['Total_Sell'].eq(0)&summary['Total_Cost'].eq(0)].copy().sort_values('MAWB')
    sell_zero_only = summary[summary['Total_Sell'].eq(0)&summary['Total_Cost'].gt(0)].copy().sort_values('Total_Cost',ascending=False)
    cost_zero_only = summary[summary['Total_Cost'].eq(0)&summary['Total_Sell'].gt(0)].copy().sort_values('Total_Sell',ascending=False)

    cc = df.groupby('Charge Code',as_index=False).agg(Total_Cost=('Cost Amount','sum'),Total_Sell=('Sell Amount','sum'),Profit=('Line Profit','sum'),Line_Count=('Charge Code','size'),MAWB_Count=('MAWB',pd.Series.nunique))
    neg_counts = df.assign(_neg=df['Line Profit'].lt(0)).groupby('Charge Code',as_index=False).agg(**{'Profit<0':('_neg','sum')})
    cc = cc.merge(neg_counts,on='Charge Code',how='left').merge(_pivot(df[['MAWB','Charge Code']],'Charge Code',flags),on='Charge Code',how='left').fillna(0).sort_values('Profit',ascending=False)

    vendor = df.groupby('Vendor',as_index=False).agg(Total_Cost=('Cost Amount','sum'),Total_Sell=('Sell Amount','sum'),Profit=('Line Profit','sum'),Line_Count=('Vendor','size'),MAWB_Count=('MAWB',pd.Series.nunique))
    vendor = vendor.merge(_pivot(df[['MAWB','Vendor']],'Vendor',flags),on='Vendor',how='left').fillna(0).sort_values('Profit',ascending=False)

    cc_mawb = df.groupby(['MAWB','Charge Code'],as_index=False).agg(Client=('Client','first'),Vendor=('Vendor','first'),Branch=('Branch','first'),Total_Cost=('Cost Amount','sum'),Total_Sell=('Sell Amount','sum'),ETA=('ETA','max'))
    cc_mawb['Profit'] = cc_mawb['Total_Sell'] - cc_mawb['Total_Cost']
    cc_mawb['Profit Margin %'] = pct(cc_mawb['Profit'],cc_mawb['Total_Sell'])
    cc_mawb['ETA Month'] = pd.to_datetime(cc_mawb['ETA'],errors='coerce').dt.to_period('M').astype(str).replace('NaT','')
    def cc_exc(r):
        if r['Profit'] >= 0:
            return False
        code, client, profit = str(r['Charge Code']).upper(), str(r['Client']).upper(), float(r['Profit'])
        if code == 'TISC':
            return profit < -10
        if client == 'WHALECBOS' and code in {'TABD','DSTOR','TISC'}:
            return profit < -10
        return True
    cc_neg = cc_mawb[cc_mawb.apply(cc_exc,axis=1)].copy()
    cc_neg['Exception_Type']='Profit<0'; cc_neg['Margin_Flag']='Exception'; cc_neg['Classification']='Open'
    cc_neg = cc_neg.sort_values(['Profit','MAWB','Charge Code'])

    total = len(summary); exc_count = int(summary['Margin_Flag'].eq('Exception').sum()); exempt_count = int(summary['Margin_Flag'].eq('Exempt').sum()); normal_count = int(summary['Margin_Flag'].eq('Normal').sum())
    total_cost = float(summary['Total_Cost'].sum()); total_sell = float(summary['Total_Sell'].sum()); total_profit = float(summary['Profit'].sum()); overall_pm = total_profit/total_sell if total_sell else 0
    neg_count = int(summary['Profit'].lt(0).sum()); neg_amount = float(summary.loc[summary['Profit'].lt(0),'Profit'].sum()); neg_ratio = neg_count/total if total else 0
    metrics = {
        'Total MAWB':total,'Exception Count':exc_count,'Exception %':exc_count/total if total else 0,
        'Exempt Count':exempt_count,'Exempt %':exempt_count/total if total else 0,
        'Normal Count':normal_count,'Normal %':normal_count/total if total else 0,
        'Revenue=0 Count':int(summary['Exception_Type'].eq('Revenue=0').sum()),
        'Cost=0 Count':int(summary['Exception_Type'].eq('Cost=0').sum()),
        'Cost=Sell=0 Count':int(summary['Exception_Type'].eq('Cost=Sell=0').sum()),
        'Profit<0 Count':neg_count,
        'Margin<30% Count':int(summary['Exception_Type'].eq(f'Margin<{int(low_thr*100)}%').sum()),
        'Margin>35% Count':int(summary['Exception_Type'].eq('Margin>35%').sum()),
        'Margin>80% Count':int(summary['Exception_Type'].eq(f'Margin>{int(high_thr*100)}%').sum()),
        'Total Cost':total_cost,'Total Sell':total_sell,'Total Profit':total_profit,'Overall Profit Margin %':overall_pm,
    }
    kpi = _kpi(metrics,{'Exception %','Exempt %','Normal %','Overall Profit Margin %'})
    neg_summary = pd.DataFrame([{'Metric':'Profit < 0 Count','Value':neg_count},{'Metric':'Profit < 0 Total Amount','Value':neg_amount},{'Metric':'Profit < 0 % of MAWBs','Value':format_pct_str(neg_ratio)}])

    return AuditResult(mawb_keep,not_found,not_found_df,eta_note,kpi,neg_summary,df,summary,exceptions,client_summary,margin_outliers,negative_profit,zero_margin,zero_profit,both_zero,sell_zero_only,cost_zero_only,cc,vendor,cc_neg,margin_label,
        display_df(summary,['ETA']),display_df(exceptions,['ETA']),display_df(client_summary,['Latest_ETA']),display_df(margin_outliers,['ETA']),display_df(negative_profit,['ETA']),display_df(zero_margin,['ETA']),display_df(zero_profit,['ETA']),display_df(both_zero,['ETA']),display_df(sell_zero_only,['ETA']),display_df(cost_zero_only,['ETA']),display_df(cc),display_df(vendor),display_df(cc_neg,['ETA']))
