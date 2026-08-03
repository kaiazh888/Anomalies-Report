# MAWB Audit Analyzer V5

Files stay in the same top-level folder. Run:

```bash
pip install -r requirements.txt
streamlit run app.py
```

V5 includes:
- PROCARESX 777 consolidation priority: 125 > 932 > 001.
- The same consolidation applies to every tab, including ChargeCode Profit<0.
- HANCAIWUX DTRF/TISC/TABD/DSTOR/WIO AR allocation by AP share within each MAWB.
- Profit<0 has the highest exception priority.
- CFO Summary page contains KPI, Negative KPI, Charge Code Summary and Vendor Summary.
- MAWB Summary is an independent tab.
