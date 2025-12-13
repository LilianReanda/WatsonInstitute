import pandas as pd
from datetime import datetime

# Paths
ruta_entrada = r"C:\Users\lilia\PycharmProjects\WatsonInstitute\wellsfargo-filtros.csv"

fecha = datetime.today().strftime("%m-%d-%Y")
nombre_archivo = f"{fecha} - Wells Fargo S26 - Partial Entries Report.xlsx"
ruta_salida = rf"C:\Users\lilia\PycharmProjects\WatsonInstitute\{nombre_archivo}"

# Load CSV
df = pd.read_csv(ruta_entrada)

# Clean column names
df.columns = df.columns.str.strip()
email_col = "Email (Enter Email)"
progress_col = "Progress"

# Normalize emails: lowercase, remove spaces
df[email_col] = df[email_col].astype(str).str.strip().str.replace(r"\s+", "", regex=True).str.lower()

# Replace invalid email values with NA and remove those rows
invalid_emails = {"", "nan", "none", "null"}
df[email_col] = df[email_col].replace(list(invalid_emails), pd.NA)
df = df.dropna(subset=[email_col])

# Ensure Progress is numeric
df[progress_col] = pd.to_numeric(df[progress_col], errors="coerce").fillna(0)

# Containers for selected and removed rows
seleccionados = []
eliminados = []

# Group by email and apply selection rules
for email, grupo in df.groupby(email_col, sort=True):
    grupo_ord = grupo.sort_values(by=progress_col, ascending=False).reset_index(drop=True)

    # If any Progress = 100, remove all rows for that email
    if (grupo_ord[progress_col] == 100).any():
        for _, r in grupo_ord.iterrows():
            eliminados.append(r)
    else:
        # Keep only the highest-progress row
        fila_max = grupo_ord.iloc[0]
        seleccionados.append(fila_max)

        # Remaining rows go to elimination list
        for i in range(1, len(grupo_ord)):
            eliminados.append(grupo_ord.iloc[i])

# Convert to DataFrames
df_sel = pd.DataFrame(seleccionados) if seleccionados else pd.DataFrame(columns=df.columns)
df_elim = pd.DataFrame(columns=df.columns)


# Restore original column order
df_sel = df_sel[df.columns.tolist()]
df_elim = df_elim[df.columns.tolist()]

# Sort retained rows by Progress descending
df_sel = df_sel.sort_values(by=progress_col, ascending=False).reset_index(drop=True)

# Compute metrics
below_69 = df_sel[df_sel[progress_col] <= 69]
above_70 = df_sel[df_sel[progress_col] >= 70]

# Create summary sheet
df_resumen = pd.DataFrame({
    "Metric": ["Partial Entries", "Above 70%", "Below 69%"],
    "Count": [len(above_70) + len(below_69), len(above_70), len(below_69)]
})

# Save to Excel (Final, Summary)
with pd.ExcelWriter(ruta_salida) as writer:
    df_sel.to_excel(writer, sheet_name="Final", index=False)
    df_resumen.to_excel(writer, sheet_name="Summary", index=False)

print(f"\nTotal: {len(above_70) + len(below_69)}")
print(f"Above 70%: {len(above_70)}")
print(f"Below 69%: {len(below_69)}")