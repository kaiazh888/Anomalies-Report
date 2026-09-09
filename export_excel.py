from __future__ import annotations

import io
import pandas as pd


def _safe_len(v):

    if v is None:
        return 0

    if not isinstance(
        v,
        (
            list,
            dict,
            set
        )
    ):

        try:
            if pd.isna(v):
                return 0
        except Exception:
            pass

    return len(
        str(v)
    )


def _write_df(
    writer,
    name,
    df
):

    df.to_excel(
        writer,
        sheet_name=name,
        index=False
    )

    ws = writer.sheets[
        name
    ]

    ws.freeze_panes(
        1,
        0
    )

    if len(df.columns):

        ws.autofilter(
            0,
            0,
            max(
                len(df),
                1
            ),
            len(df.columns) - 1
        )

    for i, col in enumerate(
        df.columns
    ):

        sample = (
            df[col]
            .head(200)
            .tolist()
        )

        max_len = max(
            [len(str(col))]
            +
            [
                _safe_len(v)
                for v in sample
            ]
        )

        width = min(
            max_len + 2,
            40
        )

        ws.set_column(
            i,
            i,
            max(
                width,
                10
            )
        )


def export_to_excel(
    result
) -> bytes:

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter"
    ) as writer:

        wb = writer.book

        # ====================================================
        # SUMMARY
        # ====================================================

        ws = wb.add_worksheet(
            "Summary"
        )

        writer.sheets[
            "Summary"
        ] = ws

        title = wb.add_format(
            {
                "bold": True,
                "font_size": 14
            }
        )

        section = wb.add_format(
            {
                "bold": True,
                "font_size": 11
            }
        )

        ws.write(
            0,
            0,
            "MAWB Audit Executive Summary",
            title
        )

        row = 2

        summary_sections = [
            (
                "KPI",
                result.kpi_vertical
            ),
            (
                "Negative KPI",
                result.neg_summary
            ),
            (
                "Charge Code Summary",
                result.chargecode_summary
            ),
            (
                "Vendor Summary",
                result.vendor_summary
            )
        ]

        for label, frame in summary_sections:

            ws.write(
                row,
                0,
                label,
                section
            )

            frame.to_excel(
                writer,
                sheet_name="Summary",
                startrow=row + 1,
                index=False
            )

            row += (
                len(frame)
                + 3
            )

        ws.set_column(
            0,
            15,
            18
        )

        # ====================================================
        # DETAIL SHEETS
        # ====================================================

        sheets = {

            "Raw_Data":
                result.df,

            "Exceptions":
                result.exceptions,

            "MAWB_Summary":
                result.summary,

            "Client_Summary":
                result.client_summary,

            "Margin_Outliers":
                result.margin_outliers,

            "Negative_Profit":
                result.negative_profit,

            "Zero_Margin":
                result.zero_margin,

            "Zero_Profit":
                result.zero_profit,

            "Cost=Sell=0":
                result.both_zero,

            "Sell=0":
                result.sell_zero_only,

            "Cost=0":
                result.cost_zero_only,

            "ChargeCode_Profit<0":
                result.chargecode_profit_lt0_mawb
        }

        for name, frame in sheets.items():

            _write_df(
                writer,
                name,
                frame
            )

        # ====================================================
        # MAWB NOT FOUND
        # ====================================================

        if len(
            result.mawb_not_found_df
        ):

            _write_df(
                writer,
                "MAWB_Not_Found",
                result.mawb_not_found_df
            )

    output.seek(0)

    return output.getvalue()
