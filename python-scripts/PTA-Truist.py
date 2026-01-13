import pandas as pd
from datetime import datetime
import warnings

# Suppress openpyxl default-style warning
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Generate today's date for the output filename
today_str = datetime.today().strftime("%m-%d-%y")

# File paths
partials_file = r"C:\Users\lilia\PycharmProjects\WatsonInstitute\reports\Truist Spring 26\01-09-2026 - Truist S26 - Partial Entries Report #9.xlsx"
salesforce_file = r"C:\Users\lilia\PycharmProjects\WatsonInstitute\salesforce\truist-salesforce.xlsx"
output_file = fr"C:\Users\lilia\PycharmProjects\WatsonInstitute\{today_str}-Truist-Partials-Entries-Converted-to-Applications.xlsx"

# Load Excel files
df_partials = pd.read_excel(partials_file)
df_salesforce = pd.read_excel(salesforce_file)

# Clean and standardize email fields for matching
df_partials["Email_clean"] = df_partials["Email (Enter Email)"].astype(str).str.strip().str.lower()
df_salesforce["Email_clean"] = df_salesforce["Email"].astype(str).str.strip().str.lower()

# Merge both datasets by email
merged = pd.merge(
    df_partials,
    df_salesforce,
    on="Email_clean",
    how="inner",
    suffixes=("_partial", "_sf")
)

# Select final output columns
final_df = merged[[
    "Contact ID",
    "First Name",
    "Last Name",
    "Email",
    "Application Date Submitted"
]].copy()

# Convert Application Date Submitted to datetime
final_df["Application Date Submitted"] = pd.to_datetime(
    final_df["Application Date Submitted"], errors="coerce"
)

# Sort by application date
final_df = final_df.sort_values(by="Application Date Submitted", ascending=True)

# Convert date to MM-DD-YYYY format
final_df["Application Date Submitted"] = final_df["Application Date Submitted"].dt.strftime("%m-%d-%Y")

# Save file
final_df.to_excel(output_file, index=False)

print(f"File created successfully")