import pandas as pd
from datetime import datetime

# ---------------------------------------------------------
# INPUT CSV - Wells Fargo
# ---------------------------------------------------------
ruta = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\wellsfargo.csv"

df = pd.read_csv(ruta)

# ---------------------------------------------------------
# 1) Limpiar nombres de columnas
# ---------------------------------------------------------
df.columns = df.columns.astype(str).str.strip()

# ---------------------------------------------------------
# 2) Corregir encoding en nombres
# ---------------------------------------------------------
for col in ["Name (First)", "Name (Last)"]:
    if col in df.columns:
        df[col] = df[col].apply(
            lambda x: x.encode("latin-1", "ignore").decode("utf-8", "ignore")
            if isinstance(x, str) else x
        )

# ---------------------------------------------------------
# 3) Corregir encoding en columnas de texto
# ---------------------------------------------------------
columns_to_fix = [
    "Country of Citizenship",
    "Address (City)",
    "Address (State / Province)",
    "Address (Country)",
    "In what state does your venture primarily create impact?"
]

for col in columns_to_fix:
    if col in df.columns:
        df[col] = df[col].apply(
            lambda x: x.encode("latin-1", "ignore").decode("utf-8", "ignore")
            if isinstance(x, str) else x
        )

# ---------------------------------------------------------
# 4) Normalizar Email
# ---------------------------------------------------------
email_col = "Email (Enter Email)"

if email_col not in df.columns:
    raise KeyError(f"No se encontró la columna '{email_col}'.")

df[email_col] = df[email_col].astype(str).str.strip()
df[email_col] = df[email_col].replace(
    ["", "nan", "NaN", "None", "NULL", "null"], "Null"
)

# ---------------------------------------------------------
# 5) Normalizar Progress
# ---------------------------------------------------------
progress_col = "Progress"

if progress_col not in df.columns:
    raise KeyError("No se encontró la columna 'Progress'.")

df[progress_col] = pd.to_numeric(df[progress_col], errors="coerce")
df[progress_col] = df[progress_col].fillna(100)
df.loc[df[progress_col] == 0, progress_col] = 100

# ---------------------------------------------------------
# 6) Reordenar Progress después de Name (Last)
# ---------------------------------------------------------
if "Name (Last)" in df.columns:
    cols = df.columns.tolist()
    cols.insert(cols.index("Name (Last)") + 1, cols.pop(cols.index(progress_col)))
    df = df[cols]

# ---------------------------------------------------------
# 7) Ordenar por Email y Progress
# ---------------------------------------------------------
df = df.sort_values(by=[email_col, progress_col], ascending=[True, False])

# ---------------------------------------------------------
# 8) Normalizar email para grouping
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
# 9) SELECT PARTIAL ENTRIES
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
# 10) METRICS
# ---------------------------------------------------------
below_69 = df_selected[df_selected[progress_col] <= 69]
above_70 = df_selected[df_selected[progress_col] >= 70]

df_summary = pd.DataFrame({
    "Metric": ["Partial Entries", "Above 70%", "Below 69%"],
    "Count": [len(above_70) + len(below_69), len(above_70), len(below_69)]
})

# ---------------------------------------------------------
# 11) OUTPUT FILE
# ---------------------------------------------------------
fecha = datetime.today().strftime("%m-%d-%Y")
nombre_archivo = f"{fecha} - Wells Fargo S26 - Partial Entries Report.xlsx"
ruta_salida = rf"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\{nombre_archivo}"

with pd.ExcelWriter(ruta_salida, engine="xlsxwriter") as writer:

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

    # Auto adjust column width
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
print(f"\nReport generated: {ruta_salida}")