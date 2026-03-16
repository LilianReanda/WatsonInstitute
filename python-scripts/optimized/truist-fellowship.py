import pandas as pd
from datetime import datetime

# ---------------------------------------------------------
# INPUT CSV (Gravity Forms Export)
# ---------------------------------------------------------
ruta = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\truist.csv"

df = pd.read_csv(ruta)

# ---------------------------------------------------------
# CLEAN COLUMN NAMES
# ---------------------------------------------------------
df.columns = df.columns.astype(str).str.strip()

email_col = "Email (Enter Email)"
progress_col = "Progress"

if email_col not in df.columns:
    raise KeyError(f"No se encontró la columna '{email_col}'. Columnas disponibles: {list(df.columns)}")

if progress_col not in df.columns:
    raise KeyError("No se encontró la columna 'Progress'.")

# ---------------------------------------------------------
# NORMALIZE EMAIL
# ---------------------------------------------------------
df[email_col] = df[email_col].astype(str).str.strip()
df[email_col] = df[email_col].replace(["", "nan", "NaN", "None", "NULL", "null"], "Null")

# ---------------------------------------------------------
# CLEAN PROGRESS
# ---------------------------------------------------------
df[progress_col] = pd.to_numeric(df[progress_col], errors="coerce")
df[progress_col] = df[progress_col].fillna(100)
df.loc[df[progress_col] == 0, progress_col] = 100

# ---------------------------------------------------------
# SORT EMAIL + PROGRESS
# ---------------------------------------------------------
df = df.sort_values(by=[email_col, progress_col], ascending=[True, False])

# ---------------------------------------------------------
# MOVE PROGRESS BEFORE EMAIL
# ---------------------------------------------------------
cols = list(df.columns)
cols.remove(progress_col)
email_index = cols.index(email_col)
cols.insert(email_index, progress_col)
df = df[cols]

# ---------------------------------------------------------
# NORMALIZE EMAIL AGAIN FOR GROUPING
# ---------------------------------------------------------
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

df[progress_col] = pd.to_numeric(df[progress_col], errors="coerce").fillna(0)

# ---------------------------------------------------------
# SELECT PARTIAL ENTRIES
# ---------------------------------------------------------
selected_rows = []
removed_rows = []

for email, group in df.groupby(email_col, sort=True):

    group_sorted = group.sort_values(by=progress_col, ascending=False).reset_index(drop=True)

    if (group_sorted[progress_col] == 100).any():
        for _, row in group_sorted.iterrows():
            removed_rows.append(row)

    else:
        max_row = group_sorted.iloc[0]
        selected_rows.append(max_row)

        for i in range(1, len(group_sorted)):
            removed_rows.append(group_sorted.iloc[i])

df_selected = pd.DataFrame(selected_rows) if selected_rows else pd.DataFrame(columns=df.columns)

df_selected = df_selected[df.columns.tolist()]
df_selected = df_selected.sort_values(by=progress_col, ascending=False).reset_index(drop=True)

# ---------------------------------------------------------
# METRICS
# ---------------------------------------------------------
below_69 = df_selected[df_selected[progress_col] <= 69]
above_70 = df_selected[df_selected[progress_col] >= 70]

df_summary = pd.DataFrame({
    "Metric": ["Partial Entries", "Above 70%", "Below 69%"],
    "Count": [len(above_70) + len(below_69), len(above_70), len(below_69)]
})

# ---------------------------------------------------------
# OUTPUT FILE
# ---------------------------------------------------------
date_str = datetime.today().strftime("%m-%d-%Y")

partial_file = f"{date_str} - Truist S26 - Partial Entries Report.xlsx"
partial_output_path = rf"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\{partial_file}"

with pd.ExcelWriter(partial_output_path, engine="xlsxwriter") as writer:

    df_selected.to_excel(writer, sheet_name="Final", index=False)
    df_summary.to_excel(writer, sheet_name="Summary", index=False)

    worksheet_final = writer.sheets["Final"]
    worksheet_summary = writer.sheets["Summary"]

    # Freeze header
    worksheet_final.freeze_panes(1, 0)

    # Add filters
    worksheet_final.autofilter(
        0,
        0,
        len(df_selected),
        len(df_selected.columns) - 1
    )

    # Auto adjust column width (safe version)
    for i, col in enumerate(df_selected.columns):
        max_length = max(
            df_selected[col].astype(str).fillna("").map(len).max(),
            len(col)
        )
        worksheet_final.set_column(i, i, max_length + 2)

    for i, col in enumerate(df_summary.columns):
        max_length = max(
            df_summary[col].astype(str).fillna("").map(len).max(),
            len(col)
        )
        worksheet_summary.set_column(i, i, max_length + 2)

# ---------------------------------------------------------
# CONSOLE OUTPUT
# ---------------------------------------------------------
print(f"\nTotal Partial Entries: {len(above_70) + len(below_69)}")
print(f"Above 70%: {len(above_70)}")
print(f"Below 69%: {len(below_69)}")
print(f"\nReport generated: {partial_output_path}")