"""Generates randomized Source (CSV) and Target (Excel) datasets with random
headers, shuffled column orders, and realistic discrepancies for testing."""

import datetime
import pathlib
import random
import pandas as pd


VENDORS = [
    "Pacific Office Supplies",
    "Harbour Cloud IT Ltd",
    "Kowloon Express Courier",
    "Central Pantry Services",
    "Apex Printing House",
    "Metro Cleaning Co",
    "JobsDB Hong Kong",
    "HK Telecom Business",
    "Star Event Logistics",
    "Victoria Stationery Ltd",
]

STATUSES = ["Paid", "Pending", "Overdue"]

HEADER_POOLS = {
    "key": ["Invoice_ID", "Inv_No", "Ref_ID", "Record_ID", "Doc_Number"],
    "vendor": ["Vendor_Name", "Supplier", "Payee", "Company_Name"],
    "date": ["Invoice_Date", "Bill_Date", "Posting_Date", "Txn_Date"],
    "amount": ["Amount_HKD", "Total_HKD", "Invoice_Amount", "Net_Amount"],
    "status": ["Status", "Payment_Status", "Settlement_State", "Approval_Status"],
}


def random_date() -> str:
    start = datetime.date(2026, 1, 1)
    offset = random.randint(0, 300)
    return (start + datetime.timedelta(days=offset)).isoformat()


def random_amount() -> float:
    return round(random.uniform(250.0, 9500.0), 2)


def main() -> None:
    out_dir = pathlib.Path(__file__).resolve().parent
    n_base = random.randint(12, 25)
    prefix = random.choice(["INV", "PO", "BILL", "REF"])

    # 1. Pick random header names independently for Source and Target
    src_hdr = {k: random.choice(v) for k, v in HEADER_POOLS.items()}
    tgt_hdr = {k: random.choice(v) for k, v in HEADER_POOLS.items()}

    # 2. Build base matching records
    base_records = []
    for i in range(1, n_base + 1):
        base_records.append({
            "key": f"{prefix}-2026-{i:03d}",
            "vendor": random.choice(VENDORS),
            "date": random_date(),
            "amount": random_amount(),
            "status": random.choice(STATUSES),
        })

    src_records = [dict(r) for r in base_records]
    tgt_records = [dict(r) for r in base_records]

    # 3. Inject random discrepancies
    # Missing in Target (drop 1-2 rows from Target)
    for _ in range(random.randint(1, 2)):
        if len(tgt_records) > 5:
            tgt_records.pop(random.randrange(len(tgt_records)))

    # Missing in Source (add 1-2 extra rows to Target only)
    for j in range(1, random.randint(2, 3)):
        tgt_records.append({
            "key": f"{prefix}-2026-{n_base + j:03d}",
            "vendor": random.choice(VENDORS),
            "date": random_date(),
            "amount": random_amount(),
            "status": random.choice(STATUSES),
        })

    # Amount mismatches (modify 1-3 rows in Target)
    for r in random.sample(tgt_records, k=min(random.randint(1, 3), len(tgt_records))):
        delta = random.choice([-1, 1]) * round(random.uniform(50.0, 600.0), 2)
        r["amount"] = round(max(10.0, r["amount"] + delta), 2)

    # Status mismatches (flip 1-2 rows in Target)
    for r in random.sample(tgt_records, k=min(random.randint(1, 2), len(tgt_records))):
        other_statuses = [s for s in STATUSES if s != r["status"]]
        r["status"] = random.choice(other_statuses)

    # Duplicate rows (add 1-2 duplicates to Source and/or Target)
    src_records.append(dict(random.choice(src_records)))
    if random.choice([True, False]):
        tgt_records.append(dict(random.choice(tgt_records)))

    # 4. Shuffle row order and column order independently
    random.shuffle(src_records)
    random.shuffle(tgt_records)

    src_col_order = list(HEADER_POOLS.keys())
    tgt_col_order = list(HEADER_POOLS.keys())
    random.shuffle(src_col_order)
    random.shuffle(tgt_col_order)

    src_df = pd.DataFrame([
        {src_hdr[k]: r[k] for k in src_col_order} for r in src_records
    ])
    tgt_df = pd.DataFrame([
        {tgt_hdr[k]: r[k] for k in tgt_col_order} for r in tgt_records
    ])

    src_path = out_dir / "company_invoice_log.csv"
    tgt_path = out_dir / "vendor_statement.xlsx"
    src_df.to_csv(src_path, index=False)
    tgt_df.to_excel(tgt_path, index=False)

    print(f"Generated randomized sample files in {out_dir}")
    print(f"  Source CSV columns:   {list(src_df.columns)}")
    print(f"  Target Excel columns: {list(tgt_df.columns)}")
    print("\nRun reconciliation with:")
    print(
        f'python reconcile.py --source-key "{src_hdr["key"]}" --target-key "{tgt_hdr["key"]}" '
        f'--source-amount "{src_hdr["amount"]}" --target-amount "{tgt_hdr["amount"]}" '
        f'--source-status "{src_hdr["status"]}" --target-status "{tgt_hdr["status"]}"'
    )


if __name__ == "__main__":
    main()
