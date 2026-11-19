import argparse
import pathlib
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NAVY, SLATE, LIGHT_NAVY = "1B365D", "2C5282", "EBF1F6"
THIN = Side(style="thin", color="CBD5E0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HKD_FMT = '"HK$"#,##0.00'


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


def clean_amount(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace("HK$", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def validate_column(df: pd.DataFrame, col: str, label: str) -> None:
    if col not in df.columns:
        raise KeyError(
            f"Column '{col}' not found in {label}. Available columns: {list(df.columns)}"
        )


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
    validate_column(src_df, src_key, "Source")
    validate_column(src_df, src_amt, "Source")
    validate_column(tgt_df, tgt_key, "Target")
    validate_column(tgt_df, tgt_amt, "Target")

    check_status = (
        bool(src_status)
        and bool(tgt_status)
        and src_status in src_df.columns
        and tgt_status in tgt_df.columns
    )

    # Normalize into standard internal columns so any header names/order work seamlessly
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

    # 1. Flag duplicates before 1-to-1 matching
    src_dups = s[s.duplicated(subset=["Record_ID"], keep=False)].assign(File_Origin="Source")
    tgt_dups = t[t.duplicated(subset=["Record_ID"], keep=False)].assign(File_Origin="Target")
    dups = pd.concat([src_dups, tgt_dups], ignore_index=True)

    s_clean = s.drop_duplicates(subset=["Record_ID"], keep="first")
    t_clean = t.drop_duplicates(subset=["Record_ID"], keep="first")

    # 2. Outer join on normalized Record_ID
    merged = pd.merge(
        s_clean, t_clean, on="Record_ID", how="outer", suffixes=("_Source", "_Target"), indicator=True
    )

    # 3. Missing records
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

    # 4. Discrepancies vs. exact matches
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
    p = argparse.ArgumentParser(description="Reconcile two Excel/CSV files with flexible header mapping.")
    p.add_argument("--source", type=pathlib.Path, default=base / "sample_data" / "company_invoice_log.csv")
    p.add_argument("--target", type=pathlib.Path, default=base / "sample_data" / "vendor_statement.xlsx")
    p.add_argument("--source-sheet", default=0, help="Source Excel sheet name or 0-based index")
    p.add_argument("--target-sheet", default=0, help="Target Excel sheet name or 0-based index")

    # Shared header shortcuts (used if both files share the same column names)
    p.add_argument("--key", default="Invoice_ID", help="Shared ID column name")
    p.add_argument("--amount-col", default="Amount_HKD", help="Shared amount column name")
    p.add_argument("--status-col", default="Status", help="Shared status column name (optional)")

    # Independent overrides when Source and Target use different header names
    p.add_argument("--source-key", default=None, help="Source ID column name (overrides --key)")
    p.add_argument("--target-key", default=None, help="Target ID column name (overrides --key)")
    p.add_argument("--source-amount", default=None, help="Source amount column name (overrides --amount-col)")
    p.add_argument("--target-amount", default=None, help="Target amount column name (overrides --amount-col)")
    p.add_argument("--source-status", default=None, help="Source status column name (overrides --status-col)")
    p.add_argument("--target-status", default=None, help="Target status column name (overrides --status-col)")

    p.add_argument("--tolerance", type=float, default=0.01, help="Allowed rounding difference")
    p.add_argument("--output", type=pathlib.Path, default=base / "sample_output" / "reconciliation_report.xlsx")
    args = p.parse_args()

    src_df = load_table(args.source, sheet=args.source_sheet)
    tgt_df = load_table(args.target, sheet=args.target_sheet)

    res = reconcile(
        src_df=src_df,
        tgt_df=tgt_df,
        src_key=args.source_key or args.key,
        tgt_key=args.target_key or args.key,
        src_amt=args.source_amount or args.amount_col,
        tgt_amt=args.target_amount or args.amount_col,
        src_status=args.source_status or args.status_col,
        tgt_status=args.target_status or args.status_col,
        tolerance=args.tolerance,
    )

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
