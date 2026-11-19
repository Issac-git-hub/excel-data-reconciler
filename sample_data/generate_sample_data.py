import pathlib
import pandas as pd

def main():
    out_dir = pathlib.Path(__file__).resolve().parent
    source_rows = [
        ("INV-2026-001", "Pacific Office Supplies", "2026-10-02", 1850.00, "Paid"),
        ("INV-2026-002", "Harbour Cloud IT Ltd", "2026-10-05", 4200.00, "Paid"),
        ("INV-2026-003", "Kowloon Express Courier", "2026-10-08", 680.00, "Paid"),
        ("INV-2026-004", "Central Pantry Services", "2026-10-10", 2350.00, "Paid"),
        ("INV-2026-005", "Apex Printing House", "2026-10-12", 1500.00, "Pending"),
        ("INV-2026-006", "Metro Cleaning Co", "2026-10-15", 3800.00, "Paid"),
        ("INV-2026-007", "JobsDB Hong Kong", "2026-10-18", 5500.00, "Paid"),
        ("INV-2026-008", "HK Telecom Business", "2026-10-20", 1290.00, "Paid"),
        ("INV-2026-009", "Pacific Office Supplies", "2026-10-22", 940.00, "Pending"),
        ("INV-2026-010", "Harbour Cloud IT Ltd", "2026-10-25", 4200.00, "Paid"),
        ("INV-2026-010", "Harbour Cloud IT Ltd", "2026-10-25", 4200.00, "Paid"),
        ("INV-2026-011", "Star Event Logistics", "2026-10-28", 6400.00, "Pending"),
        ("INV-2026-012", "Apex Printing House", "2026-10-29", 2100.00, "Pending"),
    ]
    target_rows = [
        ("INV-2026-001", "Pacific Office Supplies", "2026-10-02", 1850.00, "Paid"),
        ("INV-2026-002", "Harbour Cloud IT Ltd", "2026-10-05", 4200.00, "Paid"),
        ("INV-2026-003", "Kowloon Express Courier", "2026-10-08", 680.00, "Paid"),
        ("INV-2026-004", "Central Pantry Services", "2026-10-10", 2350.00, "Paid"),
        ("INV-2026-005", "Apex Printing House", "2026-10-12", 1500.00, "Paid"),
        ("INV-2026-006", "Metro Cleaning Co", "2026-10-15", 3800.00, "Paid"),
        ("INV-2026-007", "JobsDB Hong Kong", "2026-10-18", 5500.00, "Paid"),
        ("INV-2026-008", "HK Telecom Business", "2026-10-20", 1290.00, "Paid"),
        ("INV-2026-009", "Pacific Office Supplies", "2026-10-22", 940.00, "Pending"),
        ("INV-2026-010", "Harbour Cloud IT Ltd", "2026-10-25", 4200.00, "Paid"),
        ("INV-2026-012", "Apex Printing House", "2026-10-29", 2450.00, "Pending"),
        ("INV-2026-013", "Kowloon Express Courier", "2026-10-30", 520.00, "Pending"),
    ]
    cols = ["Invoice_ID", "Vendor_Name", "Invoice_Date", "Amount_HKD", "Status"]
    pd.DataFrame(source_rows, columns=cols).to_csv(out_dir / "company_invoice_log.csv", index=False)
    pd.DataFrame(target_rows, columns=cols).to_excel(out_dir / "vendor_statement.xlsx", index=False)
    print(f"Sample files created in {out_dir}")

if __name__ == "__main__":
    main()
