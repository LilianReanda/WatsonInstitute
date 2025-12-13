import pandas as pd

# Required 2 weeks before the deadline and after the deadline
# Full path of the file
ruta = r"C:\Users\lilia\PycharmProjects\WatsonInstitute\truist.csv"

# Load CSV
df = pd.read_csv(ruta)

# Fix encoding issues in First/Last names
df['Name (First)'] = df['Name (First)'].apply(lambda x: x.encode('latin-1', 'ignore').decode('utf-8', 'ignore') if isinstance(x, str) else x)
df['Name (Last)'] = df['Name (Last)'].apply(lambda x: x.encode('latin-1', 'ignore').decode('utf-8', 'ignore') if isinstance(x, str) else x)

# Clean column names (remove invisible spaces)
df.columns = df.columns.str.strip()

# Columns to clean
columns_to_fix = [
    "Country of Citizenship",
    "Address (City)",
    "Address (State / Province)",
    "Address (Country)",
    "In what state does your venture primarily create impact?"
]

# Apply fix only to columns that exist in the DataFrame
for col in columns_to_fix:
    if col in df.columns:
        df[col] = df[col].apply(
            lambda x: x.encode('latin-1', 'ignore').decode('utf-8', 'ignore') if isinstance(x, str) else x
        )

# Fill empty Email cells with "Null"
email_col = "Email (Enter Email)"
if email_col in df.columns:
    df[email_col] = df[email_col].astype(str).str.strip()  # remove extra spaces
    df[email_col] = df[email_col].replace(["", "nan", "NaN", "None"], "Null")
else:
    print(f"Column '{email_col}' was not found.")

# Clean and fill 'Progress' column
if 'Progress' in df.columns:
    df['Progress'] = df['Progress'].astype(str).str.strip()
    df['Progress'] = df['Progress'].replace(["", "nan", "NaN", "None"], "100")
    df['Progress'] = df['Progress'].astype(float)  # convert to number for sorting
else:
    print("Column 'Progress' was not found.")

# Move 'Progress' column after 'Name (Last)'
if 'Progress' in df.columns and 'Name (Last)' in df.columns:
    cols = df.columns.tolist()
    cols.insert(cols.index("Name (Last)") + 1, cols.pop(cols.index("Progress")))
    df = df[cols]
else:
    print("Columns required to reorder do not exist.")

# Sort by Email and then by Progress (highest to lowest)
if email_col in df.columns and 'Progress' in df.columns:
    df = df.sort_values(by=[email_col, 'Progress'], ascending=[True, False])
else:
    print("Error: Email or Progress column not found.")

# Save final CSV
ruta_salida = r"C:\Users\lilia\PycharmProjects\WatsonInstitute\truist-filtros.csv"
df.to_csv(ruta_salida, index=False)

# Print number of records
print(f"File with {len(df)} entries.")