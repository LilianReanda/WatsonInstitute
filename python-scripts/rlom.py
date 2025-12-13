import pandas as pd


# Required 2 weeks before the deadline and after the deadline

# Full file path
ruta = r"C:\Users\lilia\PycharmProjects\WatsonInstitute\rlom.csv"

# Load CSV
df = pd.read_csv(ruta)

# Fix encoding issues in First/Last names
df['Name (First)'] = df['Name (First)'].apply(lambda x: x.encode('latin-1', 'ignore').decode('utf-8', 'ignore') if isinstance(x, str) else x)
df['Name (Last)'] = df['Name (Last)'].apply(lambda x: x.encode('latin-1', 'ignore').decode('utf-8', 'ignore') if isinstance(x, str) else x)

# Fix corrupted text in the Mexico states column
col_mexico = "In which state of Mexico does your venture create the most impact? If your impact extends beyond one state, please list any additional locations below."
if col_mexico in df.columns:
    df[col_mexico] = df[col_mexico].apply(
        lambda x: x.encode('latin-1', 'ignore').decode('utf-8', 'ignore') if isinstance(x, str) else x
    )

# Clean column names (remove invisible whitespace)
df.columns = df.columns.str.strip()

# Fill empty Email cells with "Null"
email_col = "Email (Enter Email)"
if email_col in df.columns:
    df[email_col] = df[email_col].astype(str).str.strip()
    df[email_col] = df[email_col].replace(["", "nan", "NaN", "None"], "Null")
else:
    print(f"Column '{email_col}' not found.")

# Clean and fill 'Progress' column
if 'Progress' in df.columns:
    df['Progress'] = df['Progress'].astype(str).str.strip()
    df['Progress'] = df['Progress'].replace(["", "nan", "NaN", "None"], "100")
    df['Progress'] = df['Progress'].astype(float)  # convert to numeric to sort
else:
    print("Column 'Progress' not found.")

# Move 'Progress' column right after 'Name (Last)'
if 'Progress' in df.columns and 'Name (Last)' in df.columns:
    cols = df.columns.tolist()
    cols.insert(cols.index("Name (Last)") + 1, cols.pop(cols.index("Progress")))

# Save final CSV
ruta_salida = r"C:\Users\lilia\PycharmProjects\WatsonInstitute\rlom-filtros.csv"
df.to_csv(ruta_salida, index=False)

# Print number of records
print(f"File with {len(df)} entries.")


