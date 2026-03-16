import pandas as pd

# INPUT CSV - Wells Fargo
ruta = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\wellsfargo.csv"

# Cargar CSV
df = pd.read_csv(ruta)

# --------------------------------------------------
# 1) Limpiar nombres de columnas (espacios invisibles)
# --------------------------------------------------
df.columns = df.columns.astype(str).str.strip()

# --------------------------------------------------
# 2) Corregir encoding en nombres
# --------------------------------------------------
for col in ["Name (First)", "Name (Last)"]:
    if col in df.columns:
        df[col] = df[col].apply(
            lambda x: x.encode("latin-1", "ignore").decode("utf-8", "ignore")
            if isinstance(x, str) else x
        )

# --------------------------------------------------
# 3) Corregir encoding en columnas de texto relevantes
# --------------------------------------------------
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

# --------------------------------------------------
# 4) Normalizar Email
# --------------------------------------------------
email_col = "Email (Enter Email)"
if email_col not in df.columns:
    raise KeyError(f"No se encontró la columna '{email_col}'.")

df[email_col] = df[email_col].astype(str).str.strip()
df[email_col] = df[email_col].replace(
    ["", "nan", "NaN", "None", "NULL", "null"], "Null"
)

# --------------------------------------------------
# 5) Normalizar Progress
# Vacío o 0 => 100
# --------------------------------------------------
if "Progress" not in df.columns:
    raise KeyError("No se encontró la columna 'Progress'.")

df["Progress"] = pd.to_numeric(df["Progress"], errors="coerce")
df["Progress"] = df["Progress"].fillna(100)
df.loc[df["Progress"] == 0, "Progress"] = 100

# --------------------------------------------------
# 6) Reordenar Progress después de Name (Last)
# --------------------------------------------------
if "Name (Last)" in df.columns:
    cols = df.columns.tolist()
    cols.insert(cols.index("Name (Last)") + 1, cols.pop(cols.index("Progress")))
    df = df[cols]

# --------------------------------------------------
# 7) Ordenar por Email y Progress (descendente)
# --------------------------------------------------
df = df.sort_values(by=[email_col, "Progress"], ascending=[True, False])

# --------------------------------------------------
# 8) Guardar CSV final
# --------------------------------------------------
ruta_salida = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\wellsfargo-filtros.csv"
df.to_csv(ruta_salida, index=False)

print(f"Script ejecutado correctamente. Filas: {len(df)}")
print(f"CSV generado: {ruta_salida}")
