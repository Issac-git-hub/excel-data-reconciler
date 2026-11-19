import argparse
import pathlib
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NAVY, SLATE = "1B365D", "2C5282"
THIN = Side(style="thin", color="CBD5E0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HKD_FMT = '"HK$"#,##0.00'

KEY_HINTS = ("id", "inv", "ref", "doc", "record", "number", "no")
KEY_EXCLUDE = ("date", "amount", "total", "hkd", "net", "price", "status", "state")
AMOUNT_HINTS = ("amount", "total", "hkd", "net", "price", "value", "cost")
STATUS_HINTS = ("status", "state", "approval", "settlement")


def load_table(path: pathlib.Path, sheet=0) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    elif path.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(path, sheet_name=sheet)
    else:
        raise ValueError(f"Unsupported file type: {path.suffix}. Use .csv or .xlsx.")

    df.columns = [str(c).strip() for c in df.columns]
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    return df


def resolve_column(
    df: pd.DataFrame,
    explicit: str | None,
    hints: tuple[str, ...],
    label: str,
    exclude: tuple[str, ...] = ("date",),
    required: bool = True,
) -> str | None:
    if explicit:
        if explicit in df.columns:
            return explicit
        raise KeyError(
            f"Column '{explicit}' not found in {label}. Available columns: {list(df.columns)}"
        )

    for col in df.columns:
        low = col.lower()
        if any(h in low for h in hints) and not any(ex in low for ex in exclude):
            return col

    if required:
        raise KeyError(
            f"Could not auto-detect column in {label}. Available columns: {list(df.columns)}"
        )
    return None


def clean_amount(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace("HK$", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def reconcile(
    src_df: pd.DataFrame,
    tgt_df: pd.DataFrame,
    src_key: str,
    tgt_key: str,
    src_amt: str,
    tgt_amt: str,
    src_status: str | None = None,
    tgt_status: str | None = None,
    tolerance: float = 0.01,
) -> dict:
    check_status = bool(src_status and tgt_status)

    s = pd.DataFrame({
        "Record_ID": src_df[src_key].astype(str).str.strip(),
        "Amount_HKD": clean_amount(src_df[src_amt]),
        "Status": src_df[src_status].astype(str).str.strip() if check_status else "",
    })
    t = pd.DataFrame({
        "Record_ID": tgt_df[tgt_key].astype(str).str.strip(),
        "Amount_HKD": clean_amount(tgt_df[tgt_amt]),
        "Status": tgt_df[tgt_status].astype(str).str.strip() if check_status else "",
    })

    src_dups = s[s.duplicated(subset=["Record_ID"], keep=False)].assign(File_Origin="Source")
    tgt_dups = t[t.duplicated(subset=["Record_ID"], keep=False)].assign(File_Origin="Target")
    dups = pd.concat([src_dups, tgt_dups], ignore_index=True)

    s_clean = s.drop_duplicates(subset=["Record_ID"], keep="first")
    t_clean = t.drop_duplicates(subset=["Record_ID"], keep="first")

    merged = pd.merge(
        s_clean, t_clean, on="Record_ID", how="outer", suffixes=("_Source", "_Target"), indicator=True
    )

    missing_rows = []
    for _, r in merged[merged["_merge"] == "left_only"].iterrows():
        missing_rows.append({
            "Record_ID": r["Record_ID"],
            "Issue": "Missing in Target",
            "Amount_HKD": r["Amount_HKD_Source"],
            "Status": r["Status_Source"],
        })
    for _, r in merged[merged["_merge"] == "right_only"].iterrows():
        missing_rows.append({
            "Record_ID": r["Record_ID"],
            "Issue": "Missing in Source",
            "Amount_HKD": r["Amount_HKD_Target"],
            "Status": r["Status_Target"],
        })

    disc_rows, match_rows = [], []
    for _, r in merged[merged["_merge"] == "both"].iterrows():
        s_val, t_val = float(r["Amount_HKD_Source"]), float(r["Amount_HKD_Target"])
        var = round(s_val - t_val, 2)
        amt_diff = abs(var) > tolerance
        st_diff = check_status and (str(r["Status_Source"]) != str(r["Status_Target"]))

        if amt_diff or st_diff:
            reasons = [lbl for lbl, flag in [("Amount Mismatch", amt_diff), ("Status Mismatch", st_diff)] if flag]
            disc_rows.append({
                "Record_ID": r["Record_ID"],
                "Discrepancy_Type": " & ".join(reasons),
                "Source_Amount_HKD": s_val,
                "Target_Amount_HKD": t_val,
                "Variance_HKD": var,
                "Source_Status": r["Status_Source"],
                "Target_Status": r["Status_Target"],
            })
        else:
            match_rows.append({
                "Record_ID": r["Record_ID"],
                "Amount_HKD": s_val,
                "Status": r["Status_Source"],
                "Result": "Matched",
            })

    return {
        "discrepancies": pd.DataFrame(disc_rows),
        "missing": pd.DataFrame(missing_rows),
        "duplicates": dups,
        "matched": pd.DataFrame(match_rows),
        "src_amt": float(s_clean["Amount_HKD"].sum()),
        "tgt_amt": float(t_clean["Amount_HKD"].sum()),
    }


def write_sheet(ws, title: str, df: pd.DataFrame) -> None:
    ws.views.sheetView[0].showGridLines = True
    cols = max(len(df.columns), 1)
    ws.merge_cells(f"A1:{get_column_letter(cols)}1")
    ws["A1"] = title
    ws["A1"].font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
    ws.row_dimensions[1].height = 26

    for idx, col in enumerate(df.columns, 1):
        c = ws.cell(row=3, column=idx, value=col.replace("_", " "))
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=SLATE, end_color=SLATE)
        c.border = BORDER

    for r_idx, (_, row) in enumerate(df.iterrows(), 4):
        for c_idx, col in enumerate(df.columns, 1):
            c = ws.cell(row=r_idx, column=c_idx, value=row[col])
            c.border = BORDER
            if "HKD" in col:
                c.number_format = HKD_FMT
                c.alignment = Alignment(horizontal="right")

    for idx, col in enumerate(df.columns, 1):
        max_len = max([len(col)] + [len(str(v)) for v in df[col]]) if len(df) else len(col)
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 5, 16), 36)
    ws.freeze_panes = "A4"


def main() -> None:
    base = pathlib.Path(__file__).resolve().parent
    p = argparse.ArgumentParser(description="Reconcile two Excel/CSV files with auto-detection or custom header mapping.")
    p.add_argument("--source", type=pathlib.Path, default=base / "sample_data" / "company_invoice_log.csv")
    p.add_argument("--target", type=pathlib.Path, default=base / "sample_data" / "vendor_statement.xlsx")
    p.add_argument("--source-sheet", default=0)
    p.add_argument("--target-sheet", default=0)
    p.add_argument("--key", default=None)
    p.add_argument("--amount-col", default=None)
    p.add_argument("--status-col", default=None)
    p.add_argument("--source-key", default=None)
    p.add_argument("--target-key", default=None)
    p.add_argument("--source-amount", default=None)
    p.add_argument("--target-amount", default=None)
    p.add_argument("--source-status", default=None)
    p.add_argument("--target-status", default=None)
    p.add_argument("--tolerance", type=float, default=0.01)
    p.add_argument("--output", type=pathlib.Path, default=base / "sample_output" / "reconciliation_report.xlsx")
    args = p.parse_args()

    src_df = load_table(args.source, sheet=args.source_sheet)
    tgt_df = load_table(args.target, sheet=args.target_sheet)

    src_key = resolve_column(src_df, args.source_key or args.key, KEY_HINTS, "Source", exclude=KEY_EXCLUDE)
    tgt_key = resolve_column(tgt_df, args.target_key or args.key, KEY_HINTS, "Target", exclude=KEY_EXCLUDE)
    src_amt = resolve_column(src_df, args.source_amount or args.amount_col, AMOUNT_HINTS, "Source")
    tgt_amt = resolve_column(tgt_df, args.target_amount or args.amount_col, AMOUNT_HINTS, "Target")
    src_st = resolve_column(src_df, args.source_status or args.status_col, STATUS_HINTS, "Source", required=False)
    tgt_st = resolve_column(tgt_df, args.target_status or args.status_col, STATUS_HINTS, "Target", required=False)

    print(f"Matched columns -> Key: [{src_key} vs {tgt_key}] | Amount: [{src_amt} vs {tgt_amt}] | Status: [{src_st} vs {tgt_st}]")

    res = reconcile(src_df, tgt_df, src_key, tgt_key, src_amt, tgt_amt, src_st, tgt_st, tolerance=args.tolerance)

    wb = openpyxl.Workbook()
    summary_df = pd.DataFrame([
        ("Exact Matches", len(res["matched"])),
        ("Discrepancies", len(res["discrepancies"])),
        ("Missing Records", len(res["missing"])),
        ("Duplicate Rows", len(res["duplicates"])),
        ("Source Total (HKD)", res["src_amt"]),
        ("Target Total (HKD)", res["tgt_amt"]),
        ("Net Variance (HKD)", round(res["src_amt"] - res["tgt_amt"], 2)),
    ], columns=["Metric", "Value"])

    write_sheet(wb.active, "Reconciliation Summary", summary_df)
    wb.active.title = "Summary"
    write_sheet(wb.create_sheet("Discrepancies"), "Amount & Status Discrepancies", res["discrepancies"])
    write_sheet(wb.create_sheet("Missing_Records"), "Missing Records", res["missing"])
    write_sheet(wb.create_sheet("Duplicates"), "Duplicate Key IDs", res["duplicates"])
    write_sheet(wb.create_sheet("Matched_Records"), "Matched Records", res["matched"])

    args.output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.output)
    print(f"Saved report to {args.output}")


if __name__ == "__main__":
    main()
