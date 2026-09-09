from __future__ import annotations

from dataclasses import dataclass
import pandas as pd

from helpers import (
    safe_numeric,
    normalize_mawb,
    parse_mawb_list,
    pct,
    clean_eta_series,
    display_df
)


# ============================================================
# BUSINESS RULES
# ============================================================

HANCAI_ALLOC_CODES = {
    "DTRF",
    "TISC",
    "TABD",
    "DSTOR",
    "WIO"
}

HANCAI_THAWB_ONLY_CODES = {
    "THAWB",
    "DDOC"
}

HANCAI_SPECIAL_CLIENTS = {
    "HANCAIWUX",
    "4PXDIGHKG"
}

SHELIU_SPECIAL_CLIENTS = {
    "SHELIUSZX",
    "LIBEXPLHR"
}


# ============================================================
# RESULT OBJECT
# ============================================================

@dataclass
class AuditResult:

    mawb_keep: list[str]
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


# ============================================================
# CLEAN TEXT
# ============================================================

def _clean_text(
    series: pd.Series,
    default="UNKNOWN"
) -> pd.Series:

    out = (
        series
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    return out.mask(
        out.eq(""),
        default
    )


# ============================================================
# READ NEW BILLING FORMAT
# ============================================================

def _prepare_billing(file) -> pd.DataFrame:

    raw = pd.read_excel(
        file,
        sheet_name=0
    )

    required = [
        "MAWB",
        "Charge Code",
        "Cost Amount",
        "Sell Amount"
    ]

    missing = [
        c for c in required
        if c not in raw.columns
    ]

    if missing:
        raise ValueError(
            "Billing file is missing required columns: "
            + ", ".join(missing)
        )

    df = raw.copy()

    # ----------------------------
    # MAWB
    # ----------------------------

    df["MAWB"] = (
        df["MAWB"]
        .apply(normalize_mawb)
    )

    df = df[
        df["MAWB"].ne("")
    ].copy()

    # ----------------------------
    # AP / AR
    # ----------------------------

    df["Cost Amount"] = safe_numeric(
        df["Cost Amount"]
    )

    df["Sell Amount"] = safe_numeric(
        df["Sell Amount"]
    )

    # ----------------------------
    # CHARGE CODE
    # ----------------------------

    df["Charge Code"] = _clean_text(
        df["Charge Code"]
    )

    # ----------------------------
    # CLIENT
    #
    # New file has both:
    # Client
    # Debtor
    #
    # Client first.
    # If Client blank -> Debtor.
    # ----------------------------

    if "Client" in df.columns:

        client = (
            df["Client"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    else:

        client = pd.Series(
            "",
            index=df.index
        )

    if "Debtor" in df.columns:

        debtor = (
            df["Debtor"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        client = client.mask(
            client.eq(""),
            debtor
        )

    df["Client"] = _clean_text(
        client
    )

    # ----------------------------
    # VENDOR
    # ----------------------------

    if "Vendor" in df.columns:

        df["Vendor"] = _clean_text(
            df["Vendor"]
        )

    else:

        df["Vendor"] = "UNKNOWN"

    # Keep original MAWB for reference
    df["Original MAWB"] = df["MAWB"]

    return df


# ============================================================
# PROCARESX MAWB MERGE
#
# PRIORITY:
#
# 125 > 932 > 001
#
# 777 merges into existing main MAWB.
# ============================================================

def _procaresx_map(
    df: pd.DataFrame
) -> dict[str, str]:

    pro = (
        df.loc[
            df["Client"].eq("PROCARESX"),
            "MAWB"
        ]
        .drop_duplicates()
    )

    buckets: dict[str, set[str]] = {}

    for mawb in pro:

        if (
            len(mawb) == 12
            and mawb[3] == "-"
        ):

            last8 = mawb[-8:]
            prefix = mawb[:3]

            buckets.setdefault(
                last8,
                set()
            ).add(prefix)

    mapping = {}

    for last8, prefixes in buckets.items():

        target = None

        # Priority
        for prefix in (
            "125",
            "932",
            "001"
        ):

            if prefix in prefixes:
                target = prefix
                break

        if (
            target
            and "777" in prefixes
        ):

            mapping[
                f"777-{last8}"
            ] = (
                f"{target}-{last8}"
            )

    return mapping


def _apply_procaresx_merge(
    df: pd.DataFrame
):

    out = df.copy()

    mapping = _procaresx_map(
        out
    )

    mask = (
        out["Client"]
        .eq("PROCARESX")
    )

    out.loc[
        mask,
        "MAWB"
    ] = (
        out.loc[
            mask,
            "MAWB"
        ]
        .replace(mapping)
    )

    return out, mapping


# ============================================================
# HANCAIWUX AR ALLOCATION
#
# DTRF / TISC / TABD / DSTOR / WIO
#
# AR is stored in DTRF.
#
# Allocate DTRF AR according to AP share.
# ============================================================

def _apply_hancai_allocation(
    df: pd.DataFrame
) -> pd.DataFrame:

    out = df.copy()

    out[
        "Original Sell Amount"
    ] = out["Sell Amount"]

    out[
        "AR Allocation Applied"
    ] = False

    eligible = (
        out["Client"].eq(
            "HANCAIWUX"
        )
        &
        out["Charge Code"].isin(
            HANCAI_ALLOC_CODES
        )
    )

    grouped = (
        out.loc[eligible]
        .groupby("MAWB")
        .groups
    )

    for mawb, idx in grouped.items():

        rows = out.loc[idx]

        # Must have DTRF
        if not rows[
            "Charge Code"
        ].eq("DTRF").any():

            continue

        # Total AP of:
        # DTRF/TISC/TABD/DSTOR/WIO

        total_ap = (
            rows["Cost Amount"]
            .sum()
        )

        # AR currently stored in DTRF

        dtrf_ar = (
            rows.loc[
                rows["Charge Code"].eq(
                    "DTRF"
                ),
                "Sell Amount"
            ]
            .sum()
        )

        if total_ap <= 0:
            continue

        allocation = (
            dtrf_ar
            *
            rows["Cost Amount"]
            /
            total_ap
        )

        out.loc[
            idx,
            "Sell Amount"
        ] = allocation

        out.loc[
            idx,
            "AR Allocation Applied"
        ] = True

    return out


# ============================================================
# ETA / BRANCH
# ============================================================

def _read_eta(
    file,
    pro_mapping
):

    if file is None:
        return None, None

    eta = pd.read_excel(
        file,
        sheet_name=0
    )

    if (
        "MAWB" not in eta.columns
        or "ETA" not in eta.columns
    ):

        return (
            None,
            "ETA file must contain MAWB and ETA columns."
        )

    keep = [
        "MAWB",
        "ETA"
    ]

    if "Branch" in eta.columns:
        branch_col = "Branch"

    elif "Destination" in eta.columns:
        branch_col = "Destination"

    else:
        branch_col = None

    if branch_col:
        keep.append(
            branch_col
        )

    eta = eta[
        keep
    ].copy()

    eta["MAWB"] = (
        eta["MAWB"]
        .apply(normalize_mawb)
        .replace(pro_mapping)
    )

    eta["ETA"] = clean_eta_series(
        eta["ETA"]
    )

    if branch_col:

        eta["Branch"] = (
            eta[branch_col]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    else:

        eta["Branch"] = ""

    eta = (
        eta.groupby(
            "MAWB",
            as_index=False
        )
        .agg(
            ETA=("ETA", "max"),
            Branch=("Branch", "last")
        )
    )

    return eta, None


# ============================================================
# ACTIVE CHARGE CODES
# ============================================================

def _active_codes(
    df: pd.DataFrame
):

    active = df.loc[
        (
            df["Cost Amount"].ne(0)
        )
        |
        (
            df["Sell Amount"].ne(0)
        ),
        [
            "MAWB",
            "Charge Code"
        ]
    ].drop_duplicates()

    return (
        active
        .groupby("MAWB")[
            "Charge Code"
        ]
        .agg(
            lambda x: set(x)
        )
        .to_dict()
    )


# ============================================================
# MAWB CLASSIFICATION
# ============================================================

def _classify(
    row,
    codes,
    low_thr,
    high_thr
):

    client = row["Client"]

    profit = float(
        row["Profit"]
    )

    margin = float(
        row["Profit Margin %"]
    )

    cost = float(
        row["Total_Cost"]
    )

    sell = float(
        row["Total_Sell"]
    )

    # ========================================================
    # HIGHEST PRIORITY
    # ========================================================

    if profit < 0:

        return (
            "Exception",
            "Open",
            "Profit<0"
        )

    # ========================================================
    # ZERO CONDITIONS
    # ========================================================

    if (
        cost == 0
        and sell == 0
    ):

        return (
            "Exception",
            "Open",
            "Cost=Sell=0"
        )

    if sell == 0:

        return (
            "Exception",
            "Open",
            "Revenue=0"
        )

    if cost == 0:

        return (
            "Exception",
            "Open",
            "Cost=0"
        )

    # ========================================================
    # HANCAIWUX / 4PXDIGHKG
    #
    # Only THAWB / DDOC active
    # ========================================================

    if (
        client
        in HANCAI_SPECIAL_CLIENTS
        and codes
        and codes.issubset(
            HANCAI_THAWB_ONLY_CODES
        )
    ):

        if margin < 0.85:

            return (
                "Exception",
                "Open",
                "Margin<85%"
            )

        return (
            "Exempt",
            "Closed",
            "THAWB/DDOC Margin>=85%"
        )

    # ========================================================
    # SHELIUSZX / LIBEXPLHR
    #
    # THAWB absent
    # ========================================================

    if (
        client
        in SHELIU_SPECIAL_CLIENTS
        and "THAWB" not in codes
    ):

        if margin > 0.35:

            return (
                "Exception",
                "Open",
                "Margin>35%"
            )

        return (
            "Exempt",
            "Closed",
            "No THAWB Margin<=35%"
        )

    # ========================================================
    # DEFAULT RULE
    # ========================================================

    if margin > high_thr:

        return (
            "Exception",
            "Open",
            f"Margin>{int(high_thr * 100)}%"
        )

    if margin < low_thr:

        return (
            "Exception",
            "Open",
            f"Margin<{int(low_thr * 100)}%"
        )

    return (
        "Normal",
        "Closed",
        ""
    )


# ============================================================
# MAIN AUDIT
# ============================================================

def run_audit(
    billing_file,
    eta_file=None,
    mawb_text="",
    low_thr=0.30,
    high_thr=0.80
) -> AuditResult:

    # ========================================================
    # 1. READ BILLING
    # ========================================================

    df = _prepare_billing(
        billing_file
    )

    # ========================================================
    # 2. PROCARESX MERGE
    # ========================================================

    df, pro_mapping = (
        _apply_procaresx_merge(
            df
        )
    )

    # ========================================================
    # 3. OPTIONAL MAWB FILTER
    # ========================================================

    requested = parse_mawb_list(
        mawb_text
    )

    requested = sorted({
        pro_mapping.get(
            x,
            x
        )
        for x in requested
    })

    if requested:

        df = df[
            df["MAWB"].isin(
                requested
            )
        ].copy()

        missing = sorted(
            set(requested)
            -
            set(df["MAWB"])
        )

    else:

        missing = []

    # ========================================================
    # 4. ETA / BRANCH
    # ========================================================

    eta, eta_note = _read_eta(
        eta_file,
        pro_mapping
    )

    if eta is not None:

        df = df.merge(
            eta,
            on="MAWB",
            how="left"
        )

    else:

        df["ETA"] = pd.NaT
        df["Branch"] = ""

    # ========================================================
    # 5. HANCAIWUX AR ALLOCATION
    # ========================================================

    df = _apply_hancai_allocation(
        df
    )

    # ========================================================
    # 6. LINE PROFIT
    # ========================================================

    df["Line Profit"] = (
        df["Sell Amount"]
        -
        df["Cost Amount"]
    )

    # ========================================================
    # 7. ACTIVE CHARGE CODES
    # ========================================================

    codes_by_mawb = (
        _active_codes(df)
    )

    # ========================================================
    # 8. MAWB SUMMARY
    # ========================================================

    summary = (
        df.groupby(
            "MAWB",
            as_index=False
        )
        .agg(
            Client=(
                "Client",
                "first"
            ),
            Branch=(
                "Branch",
                "first"
            ),
            Total_Cost=(
                "Cost Amount",
                "sum"
            ),
            Total_Sell=(
                "Sell Amount",
                "sum"
            ),
            Line_Count=(
                "MAWB",
                "size"
            ),
            ETA=(
                "ETA",
                "max"
            )
        )
    )

    summary["Profit"] = (
        summary["Total_Sell"]
        -
        summary["Total_Cost"]
    )

    summary[
        "Profit Margin %"
    ] = pct(
        summary["Profit"],
        summary["Total_Sell"]
    )

    # ========================================================
    # 9. CLASSIFICATION
    # ========================================================

    flags = summary.apply(
        lambda r: _classify(
            r,
            codes_by_mawb.get(
                r["MAWB"],
                set()
            ),
            low_thr,
            high_thr
        ),
        axis=1,
        result_type="expand"
    )

    flags.columns = [
        "Margin_Flag",
        "Classification",
        "Exception_Type"
    ]

    summary = pd.concat(
        [
            summary,
            flags
        ],
        axis=1
    )

    # ========================================================
    # 10. EXCEPTION TABS
    # ========================================================

    exceptions = summary[
        summary[
            "Classification"
        ].eq("Open")
    ].copy()

    negative_profit = summary[
        summary["Profit"].lt(0)
    ].copy()

    zero_margin = summary[
        summary[
            "Profit Margin %"
        ].eq(0)
    ].copy()

    zero_profit = summary[
        summary["Profit"].eq(0)
    ].copy()

    both_zero = summary[
        summary["Total_Cost"].eq(0)
        &
        summary["Total_Sell"].eq(0)
    ].copy()

    sell_zero_only = summary[
        summary["Total_Sell"].eq(0)
        &
        summary["Total_Cost"].gt(0)
    ].copy()

    cost_zero_only = summary[
        summary["Total_Cost"].eq(0)
        &
        summary["Total_Sell"].gt(0)
    ].copy()

    margin_outliers = summary[
        summary[
            "Exception_Type"
        ].str.startswith(
            "Margin",
            na=False
        )
    ].copy()

    # ========================================================
    # 11. CLIENT SUMMARY
    # ========================================================

    client_summary = (
        df.groupby(
            "Client",
            as_index=False
        )
        .agg(
            Total_Cost=(
                "Cost Amount",
                "sum"
            ),
            Total_Sell=(
                "Sell Amount",
                "sum"
            ),
            MAWB_Count=(
                "MAWB",
                pd.Series.nunique
            ),
            Line_Count=(
                "Client",
                "size"
            )
        )
    )

    client_summary["Profit"] = (
        client_summary[
            "Total_Sell"
        ]
        -
        client_summary[
            "Total_Cost"
        ]
    )

    client_summary[
        "Profit Margin %"
    ] = pct(
        client_summary["Profit"],
        client_summary[
            "Total_Sell"
        ]
    )

    # ========================================================
    # 12. CHARGE CODE SUMMARY
    # ========================================================

    chargecode_summary = (
        df.groupby(
            "Charge Code",
            as_index=False
        )
        .agg(
            Total_Cost=(
                "Cost Amount",
                "sum"
            ),
            Total_Sell=(
                "Sell Amount",
                "sum"
            ),
            Profit=(
                "Line Profit",
                "sum"
            ),
            MAWB_Count=(
                "MAWB",
                pd.Series.nunique
            ),
            Line_Count=(
                "Charge Code",
                "size"
            )
        )
    )

    neg_counts = (
        df.assign(
            _neg=df[
                "Line Profit"
            ].lt(0)
        )
        .groupby(
            "Charge Code",
            as_index=False
        )["_neg"]
        .sum()
        .rename(
            columns={
                "_neg": "Profit<0"
            }
        )
    )

    chargecode_summary = (
        chargecode_summary.merge(
            neg_counts,
            on="Charge Code",
            how="left"
        )
    )

    # ========================================================
    # 13. VENDOR SUMMARY
    # ========================================================

    vendor_summary = (
        df.groupby(
            "Vendor",
            as_index=False
        )
        .agg(
            Total_Cost=(
                "Cost Amount",
                "sum"
            ),
            Total_Sell=(
                "Sell Amount",
                "sum"
            ),
            Profit=(
                "Line Profit",
                "sum"
            ),
            MAWB_Count=(
                "MAWB",
                pd.Series.nunique
            ),
            Line_Count=(
                "Vendor",
                "size"
            )
        )
    )

    # ========================================================
    # 14. CHARGE CODE PROFIT < 0
    #
    # Uses the SAME processed DF:
    #
    # PROCARESX merge already applied
    # HANCAIWUX allocation already applied
    # ========================================================

    cc = (
        df.groupby(
            [
                "MAWB",
                "Charge Code"
            ],
            as_index=False
        )
        .agg(
            Client=(
                "Client",
                "first"
            ),
            Vendor=(
                "Vendor",
                "first"
            ),
            Branch=(
                "Branch",
                "first"
            ),
            Total_Cost=(
                "Cost Amount",
                "sum"
            ),
            Total_Sell=(
                "Sell Amount",
                "sum"
            ),
            ETA=(
                "ETA",
                "max"
            )
        )
    )

    cc["Profit"] = (
        cc["Total_Sell"]
        -
        cc["Total_Cost"]
    )

    cc[
        "Profit Margin %"
    ] = pct(
        cc["Profit"],
        cc["Total_Sell"]
    )

    # ========================================================
    # CHARGE CODE PROFIT<0 RULE
    # ========================================================

    def cc_exception(r):

        if r["Profit"] >= 0:
            return False

        code = r[
            "Charge Code"
        ]

        client = r[
            "Client"
        ]

        # TISC:
        # only profit < -10

        if code == "TISC":

            return (
                r["Profit"] < -10
            )

        # WHALECBOS:
        # TABD / DSTOR / TISC
        # only profit < -10

        if (
            client == "WHALECBOS"
            and code in {
                "TABD",
                "DSTOR",
                "TISC"
            }
        ):

            return (
                r["Profit"] < -10
            )

        # Other negative profit
        return True

    chargecode_profit_lt0_mawb = (
        cc[
            cc.apply(
                cc_exception,
                axis=1
            )
        ]
        .copy()
    )

    # Only Exception rows

    chargecode_profit_lt0_mawb[
        "Exception_Type"
    ] = "Profit<0"

    chargecode_profit_lt0_mawb[
        "Margin_Flag"
    ] = "Exception"

    chargecode_profit_lt0_mawb[
        "Classification"
    ] = "Open"

    # ========================================================
    # 15. KPI
    # ========================================================

    total_mawb = len(
        summary
    )

    exc = int(
        summary[
            "Margin_Flag"
        ].eq(
            "Exception"
        ).sum()
    )

    exempt = int(
        summary[
            "Margin_Flag"
        ].eq(
            "Exempt"
        ).sum()
    )

    normal = int(
        summary[
            "Margin_Flag"
        ].eq(
            "Normal"
        ).sum()
    )

    total_cost = float(
        summary[
            "Total_Cost"
        ].sum()
    )

    total_sell = float(
        summary[
            "Total_Sell"
        ].sum()
    )

    total_profit = float(
        summary[
            "Profit"
        ].sum()
    )

    overall_margin = (
        total_profit / total_sell
        if total_sell
        else 0
    )

    kpi_vertical = pd.DataFrame(
        [
            [
                "Total MAWB",
                total_mawb
            ],
            [
                "Exception Count",
                exc
            ],
            [
                "Exempt Count",
                exempt
            ],
            [
                "Normal Count",
                normal
            ],
            [
                "Total Cost",
                total_cost
            ],
            [
                "Total Sell",
                total_sell
            ],
            [
                "Total Profit",
                total_profit
            ],
            [
                "Overall Profit Margin %",
                f"{overall_margin:.2%}"
            ]
        ],
        columns=[
            "Metric",
            "Value"
        ]
    )

    # ========================================================
    # NEGATIVE KPI
    # ========================================================

    neg_summary = pd.DataFrame(
        [
            [
                "Profit < 0 Count",
                int(
                    summary[
                        "Profit"
                    ].lt(0).sum()
                )
            ],
            [
                "Profit < 0 Total Amount",
                float(
                    summary.loc[
                        summary[
                            "Profit"
                        ].lt(0),
                        "Profit"
                    ].sum()
                )
            ]
        ],
        columns=[
            "Metric",
            "Value"
        ]
    )

    # ========================================================
    # RETURN
    # ========================================================

    return AuditResult(

        requested,

        pd.DataFrame(
            {
                "MAWB": missing
            }
        ),

        eta_note,

        kpi_vertical,
        neg_summary,

        df,
        summary,
        exceptions,
        client_summary,
        margin_outliers,
        negative_profit,
        zero_margin,
        zero_profit,
        both_zero,
        sell_zero_only,
        cost_zero_only,
        chargecode_summary,
        vendor_summary,
        chargecode_profit_lt0_mawb,

        display_df(
            summary,
            ["ETA"]
        ),

        display_df(
            exceptions,
            ["ETA"]
        ),

        display_df(
            client_summary
        ),

        display_df(
            margin_outliers,
            ["ETA"]
        ),

        display_df(
            negative_profit,
            ["ETA"]
        ),

        display_df(
            zero_margin,
            ["ETA"]
        ),

        display_df(
            zero_profit,
            ["ETA"]
        ),

        display_df(
            both_zero,
            ["ETA"]
        ),

        display_df(
            sell_zero_only,
            ["ETA"]
        ),

        display_df(
            cost_zero_only,
            ["ETA"]
        ),

        display_df(
            chargecode_summary
        ),

        display_df(
            vendor_summary
        ),

        display_df(
            chargecode_profit_lt0_mawb,
            ["ETA"]
        )
    )
