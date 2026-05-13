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
    name = os.path.splitext(filename)[0]

    lower = name.lower()

    if "wells" in lower and "fargo" in lower:
        return "Wells Fargo"

    if "truist" in lower:
        return "Truist"

    return name

# --------------------------------------------------
# PROCESS EACH CSV
# --------------------------------------------------

for file in os.listdir(input_folder):

    if not file.endswith(".csv"):
        continue

    program_name = detect_program(file)
    input_path = os.path.join(input_folder, file)

    # --------------------------------------------------
    # READ CSV
    # --------------------------------------------------

    df = pd.read_csv(input_path)

    # --------------------------------------------------
    # CLEAN COLUMN NAMES
    # --------------------------------------------------

    df.columns = df.columns.astype(str).str.strip()

    email_col = "Email (Enter Email)"
    progress_col = "Progress"

    if email_col not in df.columns or progress_col not in df.columns:
        continue

    # --------------------------------------------------
    # FIX ENCODING SAFELY
    # --------------------------------------------------

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(
                lambda x: (
                    x.encode("latin-1", "ignore").decode("utf-8", "ignore")
                    if isinstance(x, str)
                    else x
                )
            )

    # --------------------------------------------------
    # NORMALIZE EMAIL
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
    # NORMALIZE PROGRESS
    # --------------------------------------------------

    df[progress_col] = pd.to_numeric(df[progress_col], errors="coerce")

    df[progress_col] = df[progress_col].fillna(100)

    df.loc[df[progress_col] == 0, progress_col] = 100

    # --------------------------------------------------
    # REORDER PROGRESS AFTER LAST NAME
    # --------------------------------------------------

    if "Name (Last)" in df.columns:

        cols = df.columns.tolist()

        if progress_col in cols:
            cols.insert(
                cols.index("Name (Last)") + 1,
                cols.pop(cols.index(progress_col))
            )

        df = df[cols]

    # --------------------------------------------------
    # SORT BY EMAIL + PROGRESS
    # --------------------------------------------------

    df = df.sort_values(
        by=[email_col, progress_col],
        ascending=[True, False]
    )

    # --------------------------------------------------
    # IDENTIFY PARTIAL ENTRIES
    # --------------------------------------------------

    selected_rows = []

    for email, group in df.groupby(email_col):

        group_sorted = group.sort_values(
            by=progress_col,
            ascending=False
        )

        # Skip if any entry reached 100%
        if (group_sorted[progress_col] == 100).any():
            continue

        # Keep highest progress partial
        selected_rows.append(group_sorted.iloc[0])

    df_final = pd.DataFrame(selected_rows)

    if df_final.empty:
        continue

    df_final = (
        df_final
        .sort_values(by=progress_col, ascending=False)
        .reset_index(drop=True)
    )

    # --------------------------------------------------
    # METRICS
    # --------------------------------------------------

    above_70 = df_final[df_final[progress_col] >= 70]

    below_69 = df_final[df_final[progress_col] <= 69]

    summary = [
        ("Partial Entries", len(df_final)),
        ("Above 70%", len(above_70)),
        ("Below 69%", len(below_69))
    ]

    summary_df = pd.DataFrame(
        summary,
        columns=["Metric", "Count"]
    )

    # --------------------------------------------------
    # LOCATION METRICS
    # --------------------------------------------------

    country_col = "Address (Country)"
    state_col = "Address (State / Province)"

    location_frames = []

    # Country counts
    if country_col in df_final.columns:

        country_counts = (
            df_final[country_col]
            .fillna("Unknown")
            .astype(str)
            .value_counts()
            .reset_index()
        )

        country_counts.columns = [
            "Address (Country)",
            "COUNTA of Address (Country)"
        ]

        location_frames.append(country_counts)

    # State counts
    if state_col in df_final.columns:

        state_counts = (
            df_final[state_col]
            .fillna("Unknown")
            .astype(str)
            .value_counts()
            .reset_index()
        )

        state_counts.columns = [
            "Address (State / Province)",
            "COUNTA of Address (State / Province)"
        ]

        location_frames.append(state_counts)

    # --------------------------------------------------
    # EXPORT EXCEL
    # --------------------------------------------------

    output_name = (
        f"{today} - {program_name} - Partial Entries Report.xlsx"
    )

    output_path = os.path.join(output_folder, output_name)

    with pd.ExcelWriter(
        output_path,
        engine="xlsxwriter"
    ) as writer:

        # --------------------------------------------------
        # FINAL SHEET
        # --------------------------------------------------

        df_final.to_excel(
            writer,
            sheet_name="Final",
            index=False
        )

        # --------------------------------------------------
        # SUMMARY SHEET
        # --------------------------------------------------

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        # --------------------------------------------------
        # LOCATION SHEET
        # --------------------------------------------------

        if location_frames:

            start_col = 0

            for table in location_frames:

                table.to_excel(
                    writer,
                    sheet_name="Location",
                    startrow=0,
                    startcol=start_col,
                    index=False
                )

                # Leave 2 empty columns between tables
                start_col += len(table.columns) + 2

        # --------------------------------------------------
        # FORMAT FINAL SHEET
        # --------------------------------------------------

        ws_final = writer.sheets["Final"]

        ws_final.freeze_panes(1, 0)

        ws_final.autofilter(
            0,
            0,
            len(df_final),
            len(df_final.columns) - 1
        )

        for i, col in enumerate(df_final.columns):

            series = df_final[col].fillna("").astype(str)

            max_len = max(
                series.map(len).max(),
                len(col)
            ) + 2

            ws_final.set_column(
                i,
                i,
                min(max_len, 50)
            )

        # --------------------------------------------------
        # FORMAT SUMMARY SHEET
        # --------------------------------------------------

        ws_summary = writer.sheets["Summary"]

        ws_summary.freeze_panes(1, 0)

        for i, col in enumerate(summary_df.columns):

            series = summary_df[col].astype(str)

            max_len = max(
                series.map(len).max(),
                len(col)
            ) + 2

            ws_summary.set_column(i, i, max_len)

        # --------------------------------------------------
        # FORMAT LOCATION SHEET
        # --------------------------------------------------

        if location_frames:

            ws_location = writer.sheets["Location"]

            ws_location.freeze_panes(1, 0)

            current_col = 0

            for table in location_frames:

                for i, col in enumerate(table.columns):

                    series = table[col].astype(str)

                    max_len = max(
                        series.map(len).max(),
                        len(col)
                    ) + 2

                    ws_location.set_column(
                        current_col + i,
                        current_col + i,
                        min(max_len, 40)
                    )

                current_col += len(table.columns) + 2

    # --------------------------------------------------
    # PRINT CLEAN SUMMARY
    # --------------------------------------------------

    title = f"Report generated: {output_name}"

    line = "=" * len(title)

    print("\n" + line)
    print(title)
    print(line)

    col1_width = 20
    col2_width = 10

    print(f"{'Metric':<{col1_width}}{'Count':>{col2_width}}")

    for metric, count in summary:
        print(f"{metric:<{col1_width}}{count:>{col2_width}}")

print("\nAll files processed.")