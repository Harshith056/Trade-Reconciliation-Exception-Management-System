import pandas as pd
import os

# =====================================================
# FILE PATHS
# =====================================================

TRADE_FILE = r"data\advanced_stock_trading_dataset.xlsx"
SETTLEMENT_FILE = r"data\Settlement_Book_Realistic.xlsx"

# =====================================================
# CHECK FILES EXIST
# =====================================================

if not os.path.exists(TRADE_FILE):
    raise FileNotFoundError(f"Trade file not found: {TRADE_FILE}")

if not os.path.exists(SETTLEMENT_FILE):
    raise FileNotFoundError(f"Settlement file not found: {SETTLEMENT_FILE}")

# =====================================================
# LOAD FILES
# =====================================================

trade_df = pd.read_excel(TRADE_FILE)
settle_df = pd.read_excel(SETTLEMENT_FILE)

print("\n========== DEBUG ==========")

print("\nTrade Columns:")
print(trade_df.columns.tolist())

print("\nSettlement Columns:")
print(settle_df.columns.tolist())

print("\nTrade Shape:", trade_df.shape)
print("Settlement Shape:", settle_df.shape)

print("\n===========================\n")

# =====================================================
# CREATE TRADE_ID IF MISSING
# =====================================================

if "Trade_ID" not in trade_df.columns:
    print("Trade_ID missing in Trade Book.")
    print("Creating Trade_ID automatically...")

    trade_df["Trade_ID"] = [
        f"T{i:06d}"
        for i in range(1, len(trade_df) + 1)
    ]

# =====================================================
# ENSURE TRADE_ID EXISTS IN SETTLEMENT
# =====================================================

if "Trade_ID" not in settle_df.columns:
    raise Exception(
        "Trade_ID column missing in Settlement Book."
    )

# =====================================================
# FIND DUPLICATES
# =====================================================

duplicate_trade_ids = settle_df[
    settle_df.duplicated(
        subset=["Trade_ID"],
        keep=False
    )
]["Trade_ID"].unique()

print(
    f"Duplicate Trade IDs Found: {len(duplicate_trade_ids)}"
)

# =====================================================
# MERGE DATA
# =====================================================

merged_df = trade_df.merge(
    settle_df,
    on="Trade_ID",
    how="left",
    suffixes=("_trade", "_settle")
)

print("Merged Shape:", merged_df.shape)

# =====================================================
# DETECT COLUMN NAMES
# =====================================================

quantity_trade_col = None
quantity_settle_col = None

for col in merged_df.columns:

    if "Volume_trade" == col:
        quantity_trade_col = col

    if "Volume_settle" == col:
        quantity_settle_col = col

price_trade_col = None
price_settle_col = None

for col in merged_df.columns:

    if "Close_trade" == col:
        price_trade_col = col

    if "Close_settle" == col:
        price_settle_col = col

print("\nDetected Columns")

print("Quantity Trade:", quantity_trade_col)
print("Quantity Settle:", quantity_settle_col)

print("Price Trade:", price_trade_col)
print("Price Settle:", price_settle_col)

# =====================================================
# RECONCILIATION LOGIC
# =====================================================

statuses = []

for _, row in merged_df.iterrows():

    trade_id = row["Trade_ID"]

    # Duplicate Trade

    if trade_id in duplicate_trade_ids:
        statuses.append("DUPLICATE_TRADE")
        continue

    # Missing Settlement

    if pd.isna(row.get("Settlement_Status")):
        statuses.append("MISSING_SETTLEMENT")
        continue

    # Settlement Failure

    if (
        str(
            row.get(
                "Settlement_Status",
                ""
            )
        ).upper()
        == "FAILED"
    ):
        statuses.append("SETTLEMENT_FAILURE")
        continue

    # Quantity Mismatch

    if (
        quantity_trade_col
        and quantity_settle_col
    ):

        if (
            row[quantity_trade_col]
            != row[quantity_settle_col]
        ):
            statuses.append("QUANTITY_MISMATCH")
            continue

    # Price Mismatch

    if (
        price_trade_col
        and price_settle_col
    ):

        if round(
            float(row[price_trade_col]),
            2
        ) != round(
            float(row[price_settle_col]),
            2
        ):

            statuses.append("PRICE_MISMATCH")
            continue

    statuses.append("MATCHED")

merged_df[
    "Reconciliation_Status"
] = statuses

# =====================================================
# REPORTS
# =====================================================

reconciliation_report = merged_df.copy()

exception_report = merged_df[
    merged_df[
        "Reconciliation_Status"
    ] != "MATCHED"
]

# =====================================================
# KPI CALCULATIONS
# =====================================================

total_trades = len(merged_df)

matched_count = (
    merged_df[
        "Reconciliation_Status"
    ] == "MATCHED"
).sum()

quantity_mismatch_count = (
    merged_df[
        "Reconciliation_Status"
    ] == "QUANTITY_MISMATCH"
).sum()

price_mismatch_count = (
    merged_df[
        "Reconciliation_Status"
    ] == "PRICE_MISMATCH"
).sum()

missing_settlement_count = (
    merged_df[
        "Reconciliation_Status"
    ] == "MISSING_SETTLEMENT"
).sum()

settlement_failure_count = (
    merged_df[
        "Reconciliation_Status"
    ] == "SETTLEMENT_FAILURE"
).sum()

duplicate_trade_count = (
    merged_df[
        "Reconciliation_Status"
    ] == "DUPLICATE_TRADE"
).sum()

exception_count = (
    total_trades
    - matched_count
)

break_rate = round(
    (
        exception_count
        / total_trades
    ) * 100,
    2
)

success_rate = round(
    (
        matched_count
        / total_trades
    ) * 100,
    2
)

kpi_df = pd.DataFrame(
    {
        "Metric": [
            "Total Trades",
            "Matched Trades",
            "Total Exceptions",
            "Break Rate %",
            "Settlement Success Rate %",
            "Quantity Mismatches",
            "Price Mismatches",
            "Missing Settlements",
            "Settlement Failures",
            "Duplicate Trades"
        ],
        "Value": [
            total_trades,
            matched_count,
            exception_count,
            break_rate,
            success_rate,
            quantity_mismatch_count,
            price_mismatch_count,
            missing_settlement_count,
            settlement_failure_count,
            duplicate_trade_count
        ]
    }
)

# =====================================================
# SAVE REPORTS INTO ONE WORKBOOK
# =====================================================

output_file = "Reconciliation_System_Report.xlsx"

with pd.ExcelWriter(
    output_file,
    engine="openpyxl"
) as writer:

    reconciliation_report.to_excel(
        writer,
        sheet_name="Reconciliation_Report",
        index=False
    )

    exception_report.to_excel(
        writer,
        sheet_name="Exception_Report",
        index=False
    )

    kpi_df.to_excel(
        writer,
        sheet_name="KPI_Report",
        index=False
    )

print(
    f"\nReport Generated Successfully: {output_file}"
)


# =====================================================
# SUMMARY
# =====================================================

print("\n========== KPI SUMMARY ==========")

print(f"Total Trades: {total_trades}")
print(f"Matched Trades: {matched_count}")
print(f"Exceptions: {exception_count}")

print(
    f"Break Rate: {break_rate}%"
)

print(
    f"Settlement Success Rate: {success_rate}%"
)

print(
    f"Quantity Mismatches: {quantity_mismatch_count}"
)

print(
    f"Price Mismatches: {price_mismatch_count}"
)

print(
    f"Missing Settlements: {missing_settlement_count}"
)

print(
    f"Settlement Failures: {settlement_failure_count}"
)

print(
    f"Duplicate Trades: {duplicate_trade_count}"
)

print("\nReport Generated Successfully")

print("\nWorkbook Contents")

print("1. Reconciliation_Report Sheet")
print("2. Exception_Report Sheet")
print("3. KPI_Report Sheet")

print(
    "\nOutput File: Reconciliation_System_Report.xlsx"
)
