import pandas as pd
import os
import re
from datetime import datetime

# --------------------------------------------------
# CONFIG
# --------------------------------------------------

BASE_DIR = os.path.dirname(
    os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))
    )
)

input_folder = os.path.join(
    BASE_DIR,
    "gravity-forms-csvs"
)

output_folder = os.path.join(
    BASE_DIR,
    "reports"
)

os.makedirs(output_folder, exist_ok=True)

today = datetime.today().strftime("%m-%d-%Y")


# --------------------------------------------------
# KNOWN PROGRAMS
# --------------------------------------------------
#
# Display name -> list of keyword sequences that identify it
# inside a Gravity Forms file name. Matching ignores upper/lower
# case, spaces, hyphens and underscores, and looks for WHOLE words
# (so "ford" will not match "stanford").
#
# To add a new program, add one line here.
#
# --------------------------------------------------

KNOWN_PROGRAMS = {
    "Truist":        [("truist",)],
    "Wells Fargo":   [("wells", "fargo"), ("wellsfargo",)],
    "Google":        [("google",)],
    "Halloran":      [("halloran",)],
    "Enlight":       [("enlight",)],
    "Flagship":      [("flagship",)],
    "Western Union": [("western", "union"), ("westernunion",)],
    "Ford":          [("ford",)],
}


# --------------------------------------------------
# FUNCTION TO DETECT PROGRAM NAME
# --------------------------------------------------

def detect_program(filename):

    """
    Detect the program from the Gravity Forms file name.

    Example:
    1-2027-truist-spring-application-2026-10-05.csv -> Truist

    If no known program is found, the file name (without
    extension) is used and a warning is printed.
    """

    name = os.path.splitext(filename)[0]

    # Split into lowercase words: letters/numbers only

    words = re.findall(
        r"[a-z0-9]+",
        name.lower()
    )

    matches = []

    for program, keyword_sets in KNOWN_PROGRAMS.items():

        for keywords in keyword_sets:

            size = len(keywords)

            for i in range(len(words) - size + 1):

                if tuple(words[i:i + size]) == keywords:

                    if program not in matches:
                        matches.append(program)

    if len(matches) == 1:

        return matches[0]

    if len(matches) > 1:

        print(
            f"\nWARNING: {filename} matches several programs "
            f"({', '.join(matches)}). Using {matches[0]}."
        )

        return matches[0]

    print(
        f"\nWARNING: no known program found in {filename}. "
        f"Using the file name as the program name."
    )

    return name


# --------------------------------------------------
# FUNCTION TO READ CSV
# --------------------------------------------------

def read_gravity_forms_csv(file_path):

    """
    Read Gravity Forms CSV.

    UTF-8 with BOM is the normal expected format.
    Latin-1 is used as a fallback if needed.
    """

    try:
        return pd.read_csv(
            file_path,
            encoding="utf-8-sig"
        )

    except UnicodeDecodeError:
        return pd.read_csv(
            file_path,
            encoding="latin-1"
        )


# --------------------------------------------------
# VALIDATE INPUT FOLDER
# --------------------------------------------------

if not os.path.exists(input_folder):

    print("\nERROR:")
    print(f"Input folder not found:\n{input_folder}")

    raise FileNotFoundError(
        f"Missing folder: {input_folder}"
    )


# --------------------------------------------------
# PROCESS EACH CSV
# --------------------------------------------------

programs_seen = set()

for file in os.listdir(input_folder):

    if not file.lower().endswith(".csv"):
        continue

    program_name = detect_program(file)

    if program_name in programs_seen:

        print(
            f"\nWARNING: more than one CSV found for "
            f"{program_name} ({file}). "
            f"The report for this program will be overwritten."
        )

    programs_seen.add(program_name)

    input_path = os.path.join(
        input_folder,
        file
    )

    # --------------------------------------------------
    # READ CSV
    # --------------------------------------------------

    try:
        df = read_gravity_forms_csv(input_path)

    except Exception as e:

        print(
            f"\nCould not read file: {file}"
        )

        print(f"Error: {e}")

        continue

    # --------------------------------------------------
    # CLEAN COLUMN NAMES
    # --------------------------------------------------

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    email_col = "Email (Enter Email)"
    progress_col = "Progress"

    if (
        email_col not in df.columns
        or progress_col not in df.columns
    ):

        print(
            f"\nSkipped {file}: "
            "required columns were not found."
        )

        continue

    # --------------------------------------------------
    # NORMALIZE EMAIL
    # --------------------------------------------------

    df[email_col] = (
        df[email_col]
        .astype(str)
        .str.strip()
        .str.replace(
            r"\s+",
            "",
            regex=True
        )
        .str.lower()
    )

    invalid_emails = {
        "",
        "nan",
        "none",
        "null"
    }

    df[email_col] = df[email_col].replace(
        list(invalid_emails),
        pd.NA
    )

    df = df.dropna(
        subset=[email_col]
    )

    # --------------------------------------------------
    # NORMALIZE PROGRESS
    # --------------------------------------------------

    df[progress_col] = pd.to_numeric(
        df[progress_col],
        errors="coerce"
    )

    # Gravity Forms:
    # 0 = completed
    #
    # Blank / invalid progress is also treated as
    # completed, preserving the original behavior.

    df[progress_col] = (
        df[progress_col]
        .fillna(100)
    )

    df.loc[
        df[progress_col] == 0,
        progress_col
    ] = 100

    # --------------------------------------------------
    # REORDER PROGRESS AFTER LAST NAME
    # --------------------------------------------------

    if "Name (Last)" in df.columns:

        cols = df.columns.tolist()

        if progress_col in cols:

            progress_position = cols.index(
                progress_col
            )

            last_name_position = cols.index(
                "Name (Last)"
            )

            progress_column = cols.pop(
                progress_position
            )

            cols.insert(
                last_name_position + 1,
                progress_column
            )

        df = df[cols]

    # --------------------------------------------------
    # SORT BY EMAIL + PROGRESS
    # --------------------------------------------------

    df = df.sort_values(
        by=[
            email_col,
            progress_col
        ],
        ascending=[
            True,
            False
        ]
    )

    # --------------------------------------------------
    # IDENTIFY PARTIAL ENTRIES
    # --------------------------------------------------

    selected_rows = []

    for email, group in df.groupby(
        email_col,
        sort=False
    ):

        group_sorted = group.sort_values(
            by=progress_col,
            ascending=False
        )

        # If this person has ANY completed entry,
        # remove them from the partial-entry report.

        if (
            group_sorted[progress_col] == 100
        ).any():

            continue

        # Otherwise, keep the entry with the
        # highest completion percentage.

        selected_rows.append(
            group_sorted.iloc[0]
        )

    # --------------------------------------------------
    # BUILD FINAL DATAFRAME
    # --------------------------------------------------

    df_final = pd.DataFrame(
        selected_rows
    )

    if df_final.empty:

        print(
            f"\nNo partial entries found for "
            f"{program_name}."
        )

        continue

    df_final = (
        df_final
        .sort_values(
            by=progress_col,
            ascending=False
        )
        .reset_index(drop=True)
    )

    # --------------------------------------------------
    # METRICS
    # --------------------------------------------------

    above_70 = df_final[
        df_final[progress_col] >= 70
    ]

    below_69 = df_final[
        df_final[progress_col] <= 69
    ]

    summary = [
        (
            "Partial Entries",
            len(df_final)
        ),
        (
            "Above 70%",
            len(above_70)
        ),
        (
            "Below 69%",
            len(below_69)
        )
    ]

    summary_df = pd.DataFrame(
        summary,
        columns=[
            "Metric",
            "Count"
        ]
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

        location_frames.append(
            country_counts
        )

    # State / Province counts

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

        location_frames.append(
            state_counts
        )

    # --------------------------------------------------
    # EXPORT EXCEL
    # --------------------------------------------------

    output_name = (
        f"{today} - "
        f"{program_name} - "
        f"Partial Entries Report.xlsx"
    )

    output_path = os.path.join(
        output_folder,
        output_name
    )

    with pd.ExcelWriter(
        output_path,
        engine="xlsxwriter"
    ) as writer:

        # --------------------------------------------------
        # 1. FINAL
        # --------------------------------------------------

        df_final.to_excel(
            writer,
            sheet_name="Final",
            index=False
        )

        # --------------------------------------------------
        # 2. LOCATION
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

                start_col += (
                    len(table.columns) + 2
                )

        else:

            # Create the sheet even if there are
            # no location fields available.
            pd.DataFrame(
                {
                    "Message": [
                        "No location data available."
                    ]
                }
            ).to_excel(
                writer,
                sheet_name="Location",
                index=False
            )

        # --------------------------------------------------
        # 3. SUMMARY
        # --------------------------------------------------

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        # ==================================================
        # FORMAT FINAL
        # ==================================================

        ws_final = writer.sheets["Final"]

        ws_final.freeze_panes(
            1,
            0
        )

        ws_final.autofilter(
            0,
            0,
            len(df_final),
            len(df_final.columns) - 1
        )

        for i, col in enumerate(
            df_final.columns
        ):

            series = (
                df_final[col]
                .fillna("")
                .astype(str)
            )

            max_len = max(
                series.map(len).max(),
                len(col)
            ) + 2

            ws_final.set_column(
                i,
                i,
                min(max_len, 50)
            )

        # ==================================================
        # FORMAT LOCATION
        # ==================================================

        ws_location = writer.sheets["Location"]

        ws_location.freeze_panes(
            1,
            0
        )

        current_col = 0

        for table in location_frames:

            for i, col in enumerate(
                table.columns
            ):

                series = (
                    table[col]
                    .fillna("")
                    .astype(str)
                )

                max_len = max(
                    series.map(len).max(),
                    len(col)
                ) + 2

                ws_location.set_column(
                    current_col + i,
                    current_col + i,
                    min(max_len, 40)
                )

            current_col += (
                len(table.columns) + 2
            )

        # ==================================================
        # FORMAT SUMMARY
        # ==================================================

        ws_summary = writer.sheets["Summary"]

        ws_summary.freeze_panes(
            1,
            0
        )

        for i, col in enumerate(
            summary_df.columns
        ):

            series = (
                summary_df[col]
                .fillna("")
                .astype(str)
            )

            max_len = max(
                series.map(len).max(),
                len(col)
            ) + 2

            ws_summary.set_column(
                i,
                i,
                max_len
            )

    # --------------------------------------------------
    # PRINT CLEAN SUMMARY
    # --------------------------------------------------

    title = (
        f"Report generated: "
        f"{output_name}"
    )

    print(f"\n{title}")

    col1_width = 20
    col2_width = 10

    print(
        f"{'Metric':<{col1_width}}"
        f"{'Count':>{col2_width}}"
    )

    for metric, count in summary:

        print(
            f"{metric:<{col1_width}}"
            f"{count:>{col2_width}}"
        )


print("\nAll files processed.")