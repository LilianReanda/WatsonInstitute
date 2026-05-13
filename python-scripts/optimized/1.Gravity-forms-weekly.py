import pandas as pd
import os
import re
from datetime import datetime
from collections import defaultdict

# --------------------------------------------------
# CONFIG
# --------------------------------------------------

reports_folder = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\reports"

email_col = "Email (Enter Email)"
progress_col = "Progress"

# --------------------------------------------------
# FIND ALL REPORT FILES
# --------------------------------------------------

program_files = defaultdict(list)

for root, dirs, files in os.walk(reports_folder):

    for file in files:

        if (
            file.endswith(".xlsx")
            and "Partial Entries Report" in file
            and "Weekly" not in file
        ):

            match = re.match(
                r"(\d{2}-\d{2}-\d{4}) - (.+?) - Partial Entries Report",
                file
            )

            if not match:
                continue

            try:

                report_date = datetime.strptime(
                    match.group(1),
                    "%m-%d-%Y"
                )

                program_name = match.group(2).strip()

                full_path = os.path.join(root, file)

                program_files[program_name].append(
                    (report_date, full_path)
                )

            except:
                pass

# --------------------------------------------------
# PROCESS EACH PROGRAM
# --------------------------------------------------

for program, files in program_files.items():

    if len(files) < 2:
        continue

    # --------------------------------------------------
    # SORT FILES
    # --------------------------------------------------

    files.sort(key=lambda x: x[0])

    previous_date, previous_file = files[-2]
    latest_date, latest_file = files[-1]

    latest_filename = os.path.basename(latest_file)

    # --------------------------------------------------
    # LOAD FINAL SHEETS
    # --------------------------------------------------

    try:

        df_previous = pd.read_excel(
            previous_file,
            sheet_name="Final"
        )

        df_latest = pd.read_excel(
            latest_file,
            sheet_name="Final"
        )

    except Exception as e:

        print(f"\nERROR reading Final sheets: {e}")
        continue

    # --------------------------------------------------
    # CLEAN COLUMN NAMES
    # --------------------------------------------------

    df_previous.columns = (
        df_previous.columns.astype(str).str.strip()
    )

    df_latest.columns = (
        df_latest.columns.astype(str).str.strip()
    )

    if email_col not in df_previous.columns:
        print(f"\nMissing email column in previous file.")
        continue

    if email_col not in df_latest.columns:
        print(f"\nMissing email column in latest file.")
        continue

    # --------------------------------------------------
    # NORMALIZE EMAILS
    # --------------------------------------------------

    df_previous[email_col] = (
        df_previous[email_col]
        .astype(str)
        .str.strip()
        .str.lower()
    )

    df_latest[email_col] = (
        df_latest[email_col]
        .astype(str)
        .str.strip()
        .str.lower()
    )

    # --------------------------------------------------
    # FIND NEW ENTRIES
    # --------------------------------------------------

    previous_emails = set(
        df_previous[email_col]
    )

    weekly_new = df_latest[
        ~df_latest[email_col].isin(previous_emails)
    ].copy()

    # --------------------------------------------------
    # LOAD ORIGINAL SUMMARY
    # --------------------------------------------------

    try:

        original_summary = pd.read_excel(
            latest_file,
            sheet_name="Summary"
        )

    except:

        original_summary = pd.DataFrame(
            columns=["Metric", "Count"]
        )

    # --------------------------------------------------
    # BUILD FINAL SUMMARY
    # --------------------------------------------------

    summary_append = pd.DataFrame([
        ("New Partial Entries This Week", len(weekly_new)),
        ("Latest Report", latest_filename),
        ("Previous Report", os.path.basename(previous_file))
    ], columns=["Metric", "Count"])

    summary_df = pd.concat(
        [
            original_summary,
            summary_append
        ],
        ignore_index=True
    )

    # --------------------------------------------------
    # LOAD LOCATION TAB
    # --------------------------------------------------

    location_exists = False

    try:

        location_df = pd.read_excel(
            latest_file,
            sheet_name="Location",
            header=None
        )

        location_exists = True

    except:
        pass

    # --------------------------------------------------
    # OUTPUT FILE
    # --------------------------------------------------

    output_filename = latest_filename.replace(
        ".xlsx",
        " - Weekly.xlsx"
    )

    output_path = os.path.join(
        os.path.dirname(latest_file),
        output_filename
    )

    # --------------------------------------------------
    # DELETE EXISTING FILE
    # --------------------------------------------------

    if os.path.exists(output_path):

        try:
            os.remove(output_path)

        except PermissionError:

            print("\n===================================")
            print("FILE IS OPEN - CLOSE EXCEL FILE")
            print("===================================")
            print(output_path)

            continue

    # --------------------------------------------------
    # EXPORT REPORT
    # --------------------------------------------------

    try:

        with pd.ExcelWriter(
            output_path,
            engine="xlsxwriter"
        ) as writer:

            # --------------------------------------------------
            # 1. WEEKLY
            # --------------------------------------------------

            weekly_new.to_excel(
                writer,
                sheet_name="Weekly",
                index=False
            )

            # --------------------------------------------------
            # 2. FINAL
            # --------------------------------------------------

            df_latest.to_excel(
                writer,
                sheet_name="Final",
                index=False
            )

            # --------------------------------------------------
            # 3. LOCATION
            # --------------------------------------------------

            if location_exists:

                location_df.to_excel(
                    writer,
                    sheet_name="Location",
                    index=False,
                    header=False
                )

            # --------------------------------------------------
            # 4. SUMMARY
            # --------------------------------------------------

            summary_df.to_excel(
                writer,
                sheet_name="Summary",
                index=False
            )

            # --------------------------------------------------
            # FORMAT NORMAL SHEETS
            # --------------------------------------------------

            sheets = {
                "Weekly": weekly_new,
                "Final": df_latest,
                "Summary": summary_df
            }

            for sheet, dataframe in sheets.items():

                ws = writer.sheets[sheet]

                ws.freeze_panes(1, 0)

                if len(dataframe.columns) > 0:

                    ws.autofilter(
                        0,
                        0,
                        len(dataframe),
                        len(dataframe.columns) - 1
                    )

                for i, col in enumerate(dataframe.columns):

                    series = (
                        dataframe[col]
                        .fillna("")
                        .astype(str)
                    )

                    max_len = max(
                        series.map(len).max(),
                        len(col)
                    ) + 2

                    ws.set_column(
                        i,
                        i,
                        min(max_len, 50)
                    )

            # --------------------------------------------------
            # FORMAT LOCATION
            # --------------------------------------------------

            if location_exists:

                ws_location = writer.sheets["Location"]

                ws_location.freeze_panes(1, 0)

                for i in range(location_df.shape[1]):

                    series = (
                        location_df[i]
                        .fillna("")
                        .astype(str)
                    )

                    max_len = (
                        series.map(len).max()
                    ) + 2

                    ws_location.set_column(
                        i,
                        i,
                        min(max_len, 40)
                    )

    except Exception as e:

        print(f"\nERROR writing report: {e}")
        continue

    # --------------------------------------------------
    # CONSOLE OUTPUT
    # --------------------------------------------------

    partial_entries = len(df_latest)

    if progress_col in df_latest.columns:

        above_70 = len(
            df_latest[df_latest[progress_col] >= 70]
        )

        below_69 = len(
            df_latest[df_latest[progress_col] <= 69]
        )

    else:

        above_70 = 0
        below_69 = 0

    title = f"Report generated: {output_filename}"

    line = "=" * len(title)

    print("\n" + line)
    print(title)
    print(line)

    col1_width = 35
    col2_width = 10

    print(f"{'Metric':<{col1_width}}{'Count':>{col2_width}}")

    metrics = [
        ("Partial Entries", partial_entries),
        ("Above 70%", above_70),
        ("Below 69%", below_69),
        ("New Partial Entries This Week", len(weekly_new)),
        ("Latest Report", latest_filename),
        ("Previous Report", os.path.basename(previous_file))
    ]

    for metric, value in metrics:

        print(
            f"{metric:<{col1_width}}{str(value):>{col2_width}}"
        )

print("\nAll files processed.")