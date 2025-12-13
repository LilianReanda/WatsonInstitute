import pandas as pd
from datetime import datetime

# Input path
input_path = r"C:\Users\lilia\PycharmProjects\WatsonInstitute\truist-filtros.csv"

# Date for filenames
date_str = datetime.today().strftime("%m-%d-%Y")

# Output filename for partial entries
partial_file = f"{date_str} - Truist S26 - Partial Entries Report.xlsx"
partial_output_path = rf"C:\Users\lilia\PycharmProjects\WatsonInstitute\{partial_file}"

# Load CSV
df = pd.read_csv(input_path)

# Clean column names
df.columns = df.columns.str.strip()
email_col = "Email (Enter Email)"
progress_col = "Progress"

# Normalize emails: lowercase, remove spaces
df[email_col] = (
    df[email_col]
    .astype(str)
    .str.strip()
    .str.replace(r"\s+", "", regex=True)
    .str.lower()
)

# Replace invalid emails with NA and remove them
invalid_emails = {"", "nan", "none", "null"}
df[email_col] = df[email_col].replace(list(invalid_emails), pd.NA)
df = df.dropna(subset=[email_col])

# Ensure Progress is numeric
df[progress_col] = pd.to_numeric(df[progress_col], errors="coerce").fillna(0)

# ---------------------------------------------------------
# SELECT PARTIAL ENTRIES
# ---------------------------------------------------------
selected_rows = []
removed_rows = []

for email, group in df.groupby(email_col, sort=True):
    group_sorted = group.sort_values(by=progress_col, ascending=False).reset_index(drop=True)

    # If any row has Progress = 100, skip all (they are not partial)
    if (group_sorted[progress_col] == 100).any():
        for _, row in group_sorted.iterrows():
            removed_rows.append(row)
    else:
        # Keep the highest progress row
        max_row = group_sorted.iloc[0]
        selected_rows.append(max_row)

        # Remaining go to removed list
        for i in range(1, len(group_sorted)):
            removed_rows.append(group_sorted.iloc[i])

# Convert to DataFrame
df_selected = pd.DataFrame(selected_rows) if selected_rows else pd.DataFrame(columns=df.columns)

# Restore original order and sort by progress descending
df_selected = df_selected[df.columns.tolist()]
df_selected = df_selected.sort_values(by=progress_col, ascending=False).reset_index(drop=True)

# ---------------------------------------------------------
# CREATE METRICS / SUMMARY
# ---------------------------------------------------------
below_69 = df_selected[df_selected[progress_col] <= 69]
above_70 = df_selected[df_selected[progress_col] >= 70]

df_summary = pd.DataFrame({
    "Metric": ["Partial Entries", "Above 70%", "Below 69%"],
    "Count": [len(above_70) + len(below_69), len(above_70), len(below_69)]
})

# ---------------------------------------------------------
# SAVE PARTIAL ENTRIES FILE
# ---------------------------------------------------------
with pd.ExcelWriter(partial_output_path) as writer:
    df_selected.to_excel(writer, sheet_name="Final", index=False)
    df_summary.to_excel(writer, sheet_name="Summary", index=False)

# ---------------------------------------------------------
# Console output
# ---------------------------------------------------------
print(f"\nTotal Partial Entries: {len(above_70) + len(below_69)}")
print(f"Above 70%: {len(above_70)}")
print(f"Below 69%: {len(below_69)}")
