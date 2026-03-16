import pandas as pd

# INPUT CSV exportado desde Gravity Forms
ruta = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\truist.csv"

# Cargar CSV
df = pd.read_csv(ruta)

# 1) Limpiar nombres de columnas primero (espacios invisibles)
df.columns = df.columns.astype(str).str.strip()

email_col = "Email (Enter Email)"

# 2) Normalizar Email
if email_col not in df.columns:
    raise KeyError(f"No se encontró la columna '{email_col}'. Columnas disponibles: {list(df.columns)}")

df[email_col] = df[email_col].astype(str).str.strip()
df[email_col] = df[email_col].replace(["", "nan", "NaN", "None", "NULL", "null"], "Null")

# 3) Progress: vacío o 0 -> 100
if "Progress" not in df.columns:
    raise KeyError("No se encontró la columna 'Progress'.")

# Convertir a numérico (vacíos -> NaN)
df["Progress"] = pd.to_numeric(df["Progress"], errors="coerce")

# Vacíos -> 100
df["Progress"] = df["Progress"].fillna(100)

# 0 -> 100
df.loc[df["Progress"] == 0, "Progress"] = 100

# 4) Ordenar por Email y luego por Progress (mayor a menor)
df = df.sort_values(by=[email_col, "Progress"], ascending=[True, False])

# 5) Mover la columna Progress antes de Email
cols = list(df.columns)
cols.remove("Progress")
email_index = cols.index(email_col)
cols.insert(email_index, "Progress")
df = df[cols]

# OUTPUT CSV
ruta_salida_csv = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\truist-filtros.csv"
df.to_csv(ruta_salida_csv, index=False)

print(f"Script ejecutado correctamente. Filas: {len(df)}")
print(f"CSV generado: {ruta_salida_csv}")