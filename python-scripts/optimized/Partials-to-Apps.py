import pandas as pd
from datetime import datetime
import os
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Fecha de hoy
today_str = datetime.today().strftime("%m-%d-%y")

base_path = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute"

reports_path = os.path.join(base_path, "reports")
salesforce_path = os.path.join(base_path, "salesforce")
output_path = os.path.join(base_path, "outputs")

os.makedirs(output_path, exist_ok=True)


def extract_program_name(report_file_name):
    """Extrae el nombre del programa del archivo de report"""
    name = report_file_name.replace(".xlsx", "")
    parts = name.split(" - ")

    if len(parts) >= 2:
        return parts[1].strip().lower()

    return name.lower()


def find_matching_salesforce_file(report_file_name):
    """Encuentra el archivo de salesforce que coincide con el programa"""
    program_name = extract_program_name(report_file_name)

    print(f"Looking for match with: {program_name}")

    for sf_file in os.listdir(salesforce_path):
        if not sf_file.endswith(".xlsx"):
            continue

        sf_clean = sf_file.lower()

        if program_name in sf_clean:
            return os.path.join(salesforce_path, sf_file)

    return None


# 🔁 Procesar archivos dentro de /reports
for file in os.listdir(reports_path):
    if not file.endswith(".xlsx"):
        continue

    print(f"\nProcessing: {file}")

    report_file = os.path.join(reports_path, file)
    salesforce_file = find_matching_salesforce_file(file)

    if not salesforce_file:
        print("❌ No matching salesforce file found")
        continue

    print(f"📄 Salesforce match: {os.path.basename(salesforce_file)}")

    # Leer archivos
    df_partials = pd.read_excel(report_file)
    df_salesforce = pd.read_excel(salesforce_file)

    # Limpiar emails
    df_partials["Email_clean"] = df_partials["Email (Enter Email)"].astype(str).str.strip().str.lower()
    df_salesforce["Email_clean"] = df_salesforce["Email"].astype(str).str.strip().str.lower()

    # Merge
    merged = pd.merge(df_partials, df_salesforce, on="Email_clean", how="inner")

    if merged.empty:
        print("⚠️ No matches found")
        continue

    # Selección final
    final_df = merged[[
        "Contact ID",
        "First Name",
        "Last Name",
        "Email",
        "Application Date Submitted"
    ]].copy()

    # Fechas
    final_df["Application Date Submitted"] = pd.to_datetime(
        final_df["Application Date Submitted"], errors="coerce"
    )

    final_df = final_df.sort_values(by="Application Date Submitted")

    final_df["Application Date Submitted"] = final_df["Application Date Submitted"].dt.strftime("%m-%d-%Y")

    # Nombre del output
    program_name = extract_program_name(file).replace(" ", "")
    output_file = f"{today_str}-{program_name}-Converted.xlsx"

    final_df.to_excel(os.path.join(output_path, output_file), index=False)

    print(f"✅ Output created: {output_file}")

print("\nAll files processed!")