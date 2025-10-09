import os
import re
import pandas as pd
import calendar

# === Step 1: Build dynamic path ===
# Current file directory: C:\CBSL\Script\report_automations\
current_dir = os.path.dirname(os.path.abspath(__file__))

# Go two levels up to C:\CBSL
base_dir = os.path.abspath(os.path.join(current_dir, os.pardir, os.pardir))

# Construct path to C:\CBSL\Script\working\monthly
monthly_dir = os.path.join(base_dir, "Script", "working", "monthly")

# Find the dynamic subfolder inside 'monthly' (e.g., '08-01-2025(2)')
subfolders = [f for f in os.listdir(monthly_dir) if os.path.isdir(os.path.join(monthly_dir, f))]
if not subfolders:
    raise FileNotFoundError("No subfolder found inside 'monthly' directory.")
if len(subfolders) > 1:
    print("Warning: Multiple subfolders found. Using the first one found.")

target_folder = os.path.join(monthly_dir, subfolders[0])
data_dir = os.path.join(target_folder, "NBD_QF_23_C8")

# === Step 2: Identify Client Rating and Summary files ===
client_rating_file = None
summary_file = None

for f in os.listdir(data_dir):
    if f.startswith("C8-Client Rating") and f.endswith(".xlsb"):
        client_rating_file = os.path.join(data_dir, f)
    elif f.startswith("Summary") and f.endswith(".xlsb"):
        summary_file = os.path.join(data_dir, f)

if not client_rating_file:
    raise FileNotFoundError("No file found starting with 'C8-Client Rating' and ending with .xlsb")
if not summary_file:
    raise FileNotFoundError("No file found starting with 'Summary' and ending with .xlsb")

print(f"Selected Client Rating file: {client_rating_file}")
print(f"Selected Summary file: {summary_file}")

# === Step 3: Identify latest dated sheet from Client Rating file ===
xls = pd.ExcelFile(client_rating_file, engine="pyxlsb")
sheet_names = xls.sheet_names
print("Available sheets in Client Rating file:", sheet_names)

date_pattern = re.compile(r"Rating\s*-\s*(\d{2}-\d{2}-\d{4})")
sheet_dates = {}

for sheet in sheet_names:
    match = date_pattern.search(sheet)
    if match:
        date_str = match.group(1)
        date = pd.to_datetime(date_str, format="%d-%m-%Y", errors="coerce")
        if pd.notnull(date):
            sheet_dates[sheet] = date

if not sheet_dates:
    raise ValueError("No valid sheets found with date format like 'Rating -30-06-2024'")

latest_sheet = max(sheet_dates, key=lambda k: sheet_dates[k])
date = sheet_dates[latest_sheet]
print(f"Latest Rating sheet identified: {latest_sheet}")

# === Step 4: Load DataFrames ===
df_latest_data = pd.read_excel(
    client_rating_file, 
    sheet_name=latest_sheet, 
    skiprows=1,
    engine="pyxlsb"
)
print("Client Rating DataFrame loaded successfully.")

df_summary = pd.read_excel(
    summary_file, 
    sheet_name="SUMMARY", 
    skiprows=2, 
    engine="pyxlsb"
)
print("Summary DataFrame loaded successfully.")

# Remove rows where CONTRACT NO is missing
df_summary.dropna(subset=["CONTRACT NO"], inplace=True)

print("==============================================================================")
print(f"latest data with new contracts: {df_latest_data}")
print(f"Client Rating DataFrame shape (after filtering): {df_latest_data.shape}")
print("==============================================================================")

# === Step 5: Normalize column names ===
df_latest_data.columns = df_latest_data.columns.str.strip().str.replace("\n", " ").str.replace("\r", "", regex=True)
df_summary.columns = df_summary.columns.str.strip().str.replace("\n", " ").str.replace("\r", "", regex=True)

# Identify the contract column dynamically in df_latest_data
possible_contract_cols = [col for col in df_latest_data.columns if "contract" in col.lower() and "no" in col.lower()]
if not possible_contract_cols:
    raise KeyError(f"No 'Contract No' column found in Client Rating file. Columns: {df_latest_data.columns.tolist()}")
contract_col_latest = possible_contract_cols[0]
print(f"Detected contract column in latest data: {contract_col_latest}")

# Identify the contract column in df_summary
if "CONTRACT NO" in df_summary.columns:
    summary_contract_col = "CONTRACT NO"
else:
    possible_summary_cols = [col for col in df_summary.columns if "contract" in col.lower() and "no" in col.lower()]
    if not possible_summary_cols:
        raise KeyError(f"No 'CONTRACT NO' column found in Summary file. Columns: {df_summary.columns.tolist()}")
    summary_contract_col = possible_summary_cols[0]

# Normalize contract columns
df_latest_data[contract_col_latest] = df_latest_data[contract_col_latest].astype(str).str.strip()
df_summary[summary_contract_col] = df_summary[summary_contract_col].astype(str).str.strip()

# 1️⃣ Identify closed contracts (in latest_data but missing in summary)
closed_contracts = df_latest_data[~df_latest_data[contract_col_latest].isin(df_summary[summary_contract_col])]

month_name = calendar.month_name[date.month]
merged_title = f"{month_name} - {date.year}"

if not closed_contracts.empty:
    # Create header row
    header_row = pd.DataFrame([[merged_title] + [""] * (len(closed_contracts.columns) - 1)],
                              columns=closed_contracts.columns)
    df_closed_contracts = pd.concat([header_row, closed_contracts], ignore_index=True)
    
    # Remove closed contracts from latest data
    df_latest_data = df_latest_data[df_latest_data[contract_col_latest].isin(df_summary[summary_contract_col])]
    print(f"Removed {len(closed_contracts)} closed contracts from latest data.")
else:
    df_closed_contracts = pd.DataFrame(columns=df_summary.columns)
    print("No closed contracts found.")

# 2️⃣ Identify new contracts (in summary but missing in latest_data)
new_contracts = df_summary[~df_summary[summary_contract_col].isin(df_latest_data[contract_col_latest])]

if not new_contracts.empty:
    # Keep both CONTRACT NO and CLIENT NO for mapping
    new_contracts_to_append = new_contracts[[summary_contract_col, "CLIENT NO"]].copy()
    new_contracts_to_append.rename(columns={summary_contract_col: contract_col_latest}, inplace=True)

    # Add empty columns for all other columns
    for col in df_latest_data.columns:
        if col not in new_contracts_to_append.columns:
            new_contracts_to_append[col] = ""

    # Reorder to match df_latest_data
    new_contracts_to_append = new_contracts_to_append[df_latest_data.columns]

    # Append new contracts to latest data
    df_latest_data = pd.concat([df_latest_data, new_contracts_to_append], ignore_index=True)
    print(f"Appended {len(new_contracts)} new contracts to latest data.")

    # === Perform VLOOKUP-style mapping to fill Client No ===
    if "CLIENT NO" in df_summary.columns and "Client No" in df_latest_data.columns:
        # Create a lookup dictionary from df_summary
        client_lookup = dict(zip(df_summary[summary_contract_col], df_summary["CLIENT NO"]))

        # Fill missing Client No values in df_latest_data using the lookup (VLOOKUP equivalent)
        df_latest_data["Client No"] = df_latest_data.apply(
            lambda row: row["Client No"] if pd.notna(row["Client No"]) and row["Client No"] != "" 
                        else client_lookup.get(row[contract_col_latest], ""), 
            axis=1
        )
    elif "CLIENT NO" in df_summary.columns and "Client No" in df_latest_data.columns:
        # Create a lookup dictionary from df_summary
        client_lookup = dict(zip(df_summary[summary_contract_col], df_summary["CLIENT NO"]))

        # Fill missing Client No values in df_latest_data using the lookup (VLOOKUP equivalent)
        df_latest_data["Client No"] = df_latest_data.apply(
            lambda row: row["Client No"] if pd.notna(row["Client No"]) and row["Client No"] != "" 
                        else client_lookup.get(row[contract_col_latest], ""), 
            axis=1
        )

        print("Client No column updated for newly added contracts using summary data.")
    else:
        print("Warning: CLIENT NO or Client No column not found. Skipping Client No update.")
else:
    print("No new contracts to append.")


# === FINAL STEP: Output confirmation ===
print("Data processing completed.")
print(f"closed contracts: {df_closed_contracts}")
print(f"\nClosed Contracts DataFrame: {df_closed_contracts}")
print(f"summary: {df_summary}")
print(f"Summary DataFrame shape: {df_summary.shape}")
print(f"latest data with new contracts: {df_latest_data}")
print(f"Client Rating DataFrame shape (after filtering): {df_latest_data.shape}")
print("Script execution completed successfully.")
