import pandas as pd
import os
from datetime import datetime

# --------------------------------------------------
# CONFIG
# --------------------------------------------------

input_folder = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\gravity-forms-csvs"
output_folder = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\reports"
os.makedirs(output_folder, exist_ok=True)

today = datetime.today().strftime("%m-%d-%Y")

# --------------------------------------------------
# FUNCTION TO DETECT PROGRAM NAME
# --------------------------------------------------
def detect_program(filename):
    """
    Detects program name from the filename.
    If known keywords exist (Wells Fargo, Truist), returns those.
    Otherwise returns the filename without extension.
    """
    name = os.path.splitext(filename)[0]

    lower = name.lower()
    if "wells" in lower and "fargo" in lower:
        return "Wells Fargo"
    if "truist" in lower:
        return "Truist"

    # fallback: return the filename without extension
    return name

# --------------------------------------------------
# PROCESS EACH CSV
# --------------------------------------------------
for file in os.listdir(input_folder):

    if not file.endswith(".csv"):
        continue

    program_name = detect_program(file)
    input_path = os.path.join(input_folder, file)

    print(f"\nProcessing {program_name} ({file})...")

    df = pd.read_csv(input_path)

    # --------------------------------------------------
    # Clean column names
    # --------------------------------------------------
    df.columns = df.columns.astype(str).str.strip()
    email_col = "Email (Enter Email)"
    progress_col = "Progress"

    # Skip file if required columns missing
    if email_col not in df.columns or progress_col not in df.columns:
        print(f"Skipped {file} (missing required columns)")
        continue

    # --------------------------------------------------
    # Fix encoding safely
    # --------------------------------------------------
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(
                lambda x: x.encode("latin-1", "ignore").decode("utf-8", "ignore")
                if isinstance(x, str) else x
            )

    # --------------------------------------------------
    # Normalize email
    # --------------------------------------------------
    df[email_col] = (
        df[email_col]
        .astype(str)
        .str.strip()
        .str.replace(r"\s+", "", regex=True)
        .str.lower()
    )

    invalid_emails = {"", "nan", "none", "null"}
    df[email_col] = df[email_col].replace(list(invalid_emails), pd.NA)
    df = df.dropna(subset=[email_col])

    # --------------------------------------------------
    # Normalize progress
    # --------------------------------------------------
    df[progress_col] = pd.to_numeric(df[progress_col], errors="coerce")
    df[progress_col] = df[progress_col].fillna(100)
    df.loc[df[progress_col] == 0, progress_col] = 100

    # --------------------------------------------------
    # Reorder Progress after Last Name if exists
    # --------------------------------------------------
    if "Name (Last)" in df.columns:
        cols = df.columns.tolist()
        if progress_col in cols:
            cols.insert(cols.index("Name (Last)") + 1, cols.pop(cols.index(progress_col)))
        df = df[cols]

    # --------------------------------------------------
    # Sort by Email and Progress
    # --------------------------------------------------
    df = df.sort_values(by=[email_col, progress_col], ascending=[True, False])

    # --------------------------------------------------
    # Identify Partial Entries
    # --------------------------------------------------
    selected_rows = []
    for email, group in df.groupby(email_col):
        group_sorted = group.sort_values(by=progress_col, ascending=False)
        if (group_sorted[progress_col] == 100).any():
            continue
        selected_rows.append(group_sorted.iloc[0])

    df_final = pd.DataFrame(selected_rows)
    if df_final.empty:
        print(f"No partial entries found for {program_name}")
        continue

    df_final = df_final.sort_values(by=progress_col, ascending=False).reset_index(drop=True)

    # --------------------------------------------------
    # Metrics
    # --------------------------------------------------
    above_70 = df_final[df_final[progress_col] >= 70]
    below_69 = df_final[df_final[progress_col] <= 69]

    summary = pd.DataFrame({
        "Metric": ["Partial Entries", "Above 70%", "Below 69%"],
        "Count": [len(df_final), len(above_70), len(below_69)]
    })

    # --------------------------------------------------
    # Export Excel
    # --------------------------------------------------
    output_name = f"{today} - {program_name} - Partial Entries Report.xlsx"
    output_path = os.path.join(output_folder, output_name)

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:

        # Write sheets
        df_final.to_excel(writer, sheet_name="Final", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

        workbook = writer.book

        # -----------------------------
        # Adjust Final sheet
        # -----------------------------
        ws_final = writer.sheets["Final"]
        ws_final.freeze_panes(1, 0)
        ws_final.autofilter(0, 0, len(df_final), len(df_final.columns) - 1)
        for i, col in enumerate(df_final.columns):
            series = df_final[col].fillna("").astype(str)
            max_len = max(series.map(len).max(), len(col)) + 2
            ws_final.set_column(i, i, min(max_len, 50))

        # -----------------------------
        # Adjust Summary sheet
        # -----------------------------
        ws_summary = writer.sheets["Summary"]
        ws_summary.freeze_panes(1, 0)
        for i, col in enumerate(summary.columns):
            series = summary[col].fillna("").astype(str)
            max_len = max(series.map(len).max(), len(col)) + 2
            ws_summary.set_column(i, i, min(max_len, 50))

    print(f"Report generated: {output_name}")

print("\nAll files processed.")