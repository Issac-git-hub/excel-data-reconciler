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

def load_table(path: pathlib.Path) -> pd.DataFrame:
    df = pd.read_csv(path) if path.suffix.lower() == ".csv" else pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str).str.strip()
    return df

def reconcile(src_df, tgt_df, key_col, amount_col, status_col="Status", tolerance=0.01):
    src_dups = src_df[src_df.duplicated(subset=[key_col], keep=False)].assign(File_Origin="Source")
    tgt_dups = tgt_df[tgt_df.duplicated(subset=[key_col], keep=False)].assign(File_Origin="Target")
    dups = pd.concat([src_dups, tgt_dups], ignore_index=True)

    src_c = src_df.drop_duplicates(subset=[key_col], keep="first")
    tgt_c = tgt_df.drop_duplicates(subset=[key_col], keep="first")
    merged = pd.merge(src_c, tgt_c, on=key_col, how="outer", suffixes=("_Source", "_Target"), indicator=True)

    missing_rows = []
    for _, r in merged[merged["_merge"] == "left_only"].iterrows():
        missing_rows.append({key_col: r[key_col], "Issue": "Missing in Target", "Amount_HKD": r.get(f"{amount_col}_Source", 0.0), "Status": r.get(f"{status_col}_Source", "")})
    for _, r in merged[merged["_merge"] == "right_only"].iterrows():
        missing_rows.append({key_col: r[key_col], "Issue": "Missing in Source", "Amount_HKD": r.get(f"{amount_col}_Target", 0.0), "Status": r.get(f"{status_col}_Target", "")})

    disc_rows, match_rows = [], []
    for _, r in merged[merged["_merge"] == "both"].iterrows():
        s_amt, t_amt = float(r[f"{amount_col}_Source"]), float(r[f"{amount_col}_Target"])
        var = round(s_amt - t_amt, 2)
        amt_diff = abs(var) > tolerance
        st_diff = str(r.get(f"{status_col}_Source", "")) != str(r.get(f"{status_col}_Target", ""))
        if amt_diff or st_diff:
            kind = " & ".join([x for x, flag in [("Amount Mismatch", amt_diff), ("Status Mismatch", st_diff)] if flag])
            disc_rows.append({key_col: r[key_col], "Discrepancy_Type": kind, "Source_Amount_HKD": s_amt, "Target_Amount_HKD": t_amt, "Variance_HKD": var, "Source_Status": r.get(f"{status_col}_Source", ""), "Target_Status": r.get(f"{status_col}_Target", "")})
        else:
            match_rows.append({key_col: r[key_col], "Amount_HKD": s_amt, "Status": r.get(f"{status_col}_Source", ""), "Result": "Matched"})

    return {
        "discrepancies": pd.DataFrame(disc_rows),
        "missing": pd.DataFrame(missing_rows),
        "duplicates": dups,
        "matched": pd.DataFrame(match_rows),
        "src_amt": float(src_c[amount_col].sum()),
        "tgt_amt": float(tgt_c[amount_col].sum()),
    }

def write_sheet(ws, title, df):
    ws.views.sheetView[0].showGridLines = True
    cols = max(len(df.columns), 1)
    ws.merge_cells(f"A1:{get_column_letter(cols)}1")
    ws["A1"] = title
    ws["A1"].font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
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
        ws.column_dimensions[get_column_letter(idx)].width = max(len(col) + 5, 16)
    ws.freeze_panes = "A4"

def main():
    base = pathlib.Path(__file__).resolve().parent
    p = argparse.ArgumentParser()
    p.add_argument("--source", type=pathlib.Path, default=base / "sample_data" / "company_invoice_log.csv")
    p.add_argument("--target", type=pathlib.Path, default=base / "sample_data" / "vendor_statement.xlsx")
    p.add_argument("--key", default="Invoice_ID")
    p.add_argument("--amount-col", default="Amount_HKD")
    p.add_argument("--output", type=pathlib.Path, default=base / "sample_output" / "reconciliation_report.xlsx")
    args = p.parse_args()

    res = reconcile(load_table(args.source), load_table(args.target), args.key, args.amount_col)
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
