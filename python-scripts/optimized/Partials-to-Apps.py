import pandas as pd
from datetime import datetime
import os
import warnings
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ==============================
# CONFIG
# ==============================
today_str = datetime.today().strftime("%m-%d-%y")

base_path = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute"

reports_path = os.path.join(base_path, "reports")
salesforce_path = os.path.join(base_path, "salesforce")
output_path = os.path.join(base_path, "outputs")

os.makedirs(output_path, exist_ok=True)

# ==============================
# HELPERS
# ==============================

def normalize(text):
    return re.sub(r'[^a-z0-9]', '', str(text).lower())


def extract_program_name(file_name):
    name = file_name.replace(".xlsx", "")
    parts = name.split(" - ")
    return parts[1].strip() if len(parts) >= 2 else name


def find_matching_salesforce_file(report_file_name):
    program_raw = extract_program_name(report_file_name)
    program_norm = normalize(program_raw)

    print(f"Matching: {program_raw}")

    for sf_file in os.listdir(salesforce_path):
        if not sf_file.endswith(".xlsx"):
            continue

        sf_norm = normalize(sf_file)

        if program_norm in sf_norm or sf_norm in program_norm:
            print(f"Match: {sf_file}")
            return os.path.join(salesforce_path, sf_file)

    print("No match found")
    return None


def auto_adjust_columns(file_path):
    wb = load_workbook(file_path)
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter

        for cell in col:
            if cell.value:
                length = len(str(cell.value))
                if length > max_length:
                    max_length = length

            cell.alignment = Alignment(wrap_text=True)

        ws.column_dimensions[col_letter].width = max_length + 2

    wb.save(file_path)


def safe_column(df, col_name):
    if col_name not in df.columns:
        print(f"Missing column: {col_name}")
        df[col_name] = None
    return df[col_name]


# ==============================
# MAIN
# ==============================

total_converted = 0
summary = []

for file in os.listdir(reports_path):
    if not file.endswith(".xlsx"):
        continue

    print(f"\nProcessing: {file}")

    report_file = os.path.join(reports_path, file)
    salesforce_file = find_matching_salesforce_file(file)

    if not salesforce_file:
        print("Partial Entries Converted to Apps (This Week): 0")
        continue

    try:
        df_partials = pd.read_excel(report_file)
        df_salesforce = pd.read_excel(salesforce_file)
    except Exception as e:
        print(f"Error reading file: {e}")
        continue

    # CLEAN EMAILS
    df_partials["Email_clean"] = safe_column(df_partials, "Email (Enter Email)").astype(str).str.strip().str.lower()
    df_salesforce["Email_clean"] = safe_column(df_salesforce, "Email").astype(str).str.strip().str.lower()

    # MERGE
    merged = pd.merge(df_partials, df_salesforce, on="Email_clean", how="inner")

    converted_count = len(merged)
    total_converted += converted_count

    print(f"Partial Entries Converted to Apps (This Week): {converted_count}")

    if converted_count == 0:
        continue

    program_name = extract_program_name(file)

    # FINAL DF
    final_df = pd.DataFrame({
        "Contact ID": safe_column(merged, "Contact ID"),
        "First Name": safe_column(merged, "First Name"),
        "Last Name": safe_column(merged, "Last Name"),
        "Email": safe_column(merged, "Email"),
        "Application Date Submitted": safe_column(merged, "Application Date Submitted"),
    })

    # DATE FORMAT
    final_df["Application Date Submitted"] = pd.to_datetime(
        final_df["Application Date Submitted"], errors="coerce"
    )

    final_df = final_df.sort_values(by="Application Date Submitted")
    final_df["Application Date Submitted"] = final_df["Application Date Submitted"].dt.strftime("%m-%d-%Y")

    # EXPORT
    program_name_clean = normalize(program_name)
    output_file = f"{today_str}-{program_name_clean}-Converted.xlsx"
    output_full_path = os.path.join(output_path, output_file)

    try:
        final_df.to_excel(output_full_path, index=False)
        auto_adjust_columns(output_full_path)
        print(f"Created: {output_file}")

        # Guardar en summary
        summary.append((program_name, converted_count, output_file))

    except Exception as e:
        print(f"Error writing file: {e}")


# ==============================
# SUMMARY
# ==============================

print("\n==============================")
print("SUMMARY (ONLY CONVERSIONS)")
print("==============================\n")

for program, count, file in summary:
    print(program)
    print(f"Partial Entries Converted to Apps (This Week): {count}")
    print(f"Created: {file}\n")
