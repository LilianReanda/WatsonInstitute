import pandas as pd
import os
import re
from datetime import datetime

# --------------------------------------------------
# CONFIGURATION
# --------------------------------------------------
#
# This script:
#
# 1. Searches ALL report subfolders
# 2. Finds all "Partial Entries Report" files
# 3. Groups reports by fellowship/program
# 4. Builds a week-to-week dashboard
# 5. Aligns all fellowships to the same dates
# 6. Fills missing weeks with 0
# 7. Ignores dates before recruiting launch
# 8. Excludes Wells Fargo and Truist
# 9. Calculates week-over-week growth
# 10. Imports conversion counts from:
#     Partial-Entries-Converted-to-Apps
# 11. Exports formatted Excel dashboard
#
# --------------------------------------------------

reports_folder = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\reports"

# --------------------------------------------------
# CONVERSIONS FOLDER
# --------------------------------------------------

conversions_folder = (
    r"C:\Users\Emanuel\PyCharmMiscProject"
    r"\WatsonInstitute"
    r"\Partial-Entries-Converted-to-Apps"
)

# --------------------------------------------------
# OUTPUT FOLDER
# --------------------------------------------------

output_folder = os.path.join(
    reports_folder,
    "Week-to-Week-Comparison"
)

os.makedirs(output_folder, exist_ok=True)

# --------------------------------------------------
# OUTPUT FILE
# --------------------------------------------------

today = datetime.today().strftime("%m-%d-%Y")

output_path = os.path.join(
    output_folder,
    f"{today} - Week to Week Comparison.xlsx"
)

# --------------------------------------------------
# RECRUITING LAUNCH DATE
# --------------------------------------------------

launch_date = datetime.strptime(
    "03-30-2026",
    "%m-%d-%Y"
)

# --------------------------------------------------
# FILE NAME PATTERN
# --------------------------------------------------

pattern = re.compile(
    r"(\d{2}-\d{2}-\d{4}) - (.+?) - Partial Entries Report.*\.xlsx$",
    re.IGNORECASE
)

# --------------------------------------------------
# STORE FILES BY PROGRAM
# --------------------------------------------------

program_files = {}

# --------------------------------------------------
# SEARCH REPORTS RECURSIVELY
# --------------------------------------------------

for root, dirs, files_in_dir in os.walk(reports_folder):

    if "Week-to-Week-Comparison" in root:
        continue

    for file in files_in_dir:

        match = pattern.match(file)

        if not match:
            continue

        date_str, program_name = match.groups()

        normalized_program = (
            program_name
            .strip()
            .lower()
        )

        # Skip unwanted programs
        if normalized_program in [
            "truist",
            "wells fargo"
        ]:
            continue

        display_name = program_name.strip()

        try:

            file_date = datetime.strptime(
                date_str,
                "%m-%d-%Y"
            )

        except:
            continue

        if file_date < launch_date:
            continue

        full_path = os.path.join(root, file)

        if normalized_program not in program_files:

            program_files[normalized_program] = {
                "display_name": display_name,
                "files": []
            }

        program_files[normalized_program]["files"].append(
            (file_date, full_path, file)
        )

# --------------------------------------------------
# BUILD MASTER DATE LIST
# --------------------------------------------------

all_dates = set()

for data in program_files.values():

    for file_date, _, _ in data["files"]:

        all_dates.add(file_date)

# Newest → oldest
all_dates = sorted(
    all_dates,
    reverse=True
)

# --------------------------------------------------
# LOAD CONVERSION COUNTS
# --------------------------------------------------
#
# Example file:
#
# 05-11-26-enlight-Converted.xlsx
#
# We count rows inside the file and map:
#
# enlight + 05/11/2026 = 5
#
# --------------------------------------------------

conversions_lookup = {}

conversion_pattern = re.compile(
    r"(\d{2}-\d{2}-\d{2})-(.+?)-.*\.xlsx$",
    re.IGNORECASE
)

# --------------------------------------------------
# SEARCH CONVERSION FILES RECURSIVELY
# --------------------------------------------------

for root, dirs, files in os.walk(conversions_folder):

    for file in files:

        match = conversion_pattern.match(file)

        if not match:
            continue

        date_str, raw_program = match.groups()

        normalized_program = (
            raw_program
            .strip()
            .lower()
        )

        try:

            file_date = datetime.strptime(
                date_str,
                "%m-%d-%y"
            )

            formatted_date = file_date.strftime(
                "%m/%d/%Y"
            )

        except:
            continue

        full_path = os.path.join(root, file)

        try:

            df_conversion = pd.read_excel(
                full_path
            )

            # Count rows excluding header
            conversion_count = len(
                df_conversion.index
            )

        except:

            conversion_count = 0

        if normalized_program not in conversions_lookup:

            conversions_lookup[
                normalized_program
            ] = {}

        conversions_lookup[
            normalized_program
        ][formatted_date] = conversion_count

# --------------------------------------------------
# CREATE EXCEL REPORT
# --------------------------------------------------

with pd.ExcelWriter(
    output_path,
    engine="xlsxwriter"
) as writer:

    workbook = writer.book

    # --------------------------------------------------
    # CREATE WORKSHEET
    # --------------------------------------------------

    worksheet = workbook.add_worksheet(
        "Week to Week Comparison"
    )

    writer.sheets[
        "Week to Week Comparison"
    ] = worksheet

    # --------------------------------------------------
    # FORMATS
    # --------------------------------------------------

    header_format = workbook.add_format({
        "bold": True,
        "font_color": "white",
        "bg_color": "#0B3A6E",
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "font_size": 11
    })

    date_format = workbook.add_format({
        "bold": True,
        "font_color": "white",
        "bg_color": "#0B3A6E",
        "align": "center",
        "border": 1,
        "font_size": 10
    })

    # --------------------------------------------------
    # COLORS
    # --------------------------------------------------

    light_orange = "#FCE5CD"

    metric_format = workbook.add_format({
        "bg_color": light_orange,
        "border": 1,
        "font_size": 10
    })

    value_format = workbook.add_format({
        "bg_color": light_orange,
        "border": 1,
        "align": "center",
        "font_size": 10
    })

    green_metric_format = workbook.add_format({
        "bg_color": "#D9EAD3",
        "border": 1,
        "font_size": 10
    })

    green_value_format = workbook.add_format({
        "bg_color": "#D9EAD3",
        "border": 1,
        "align": "center",
        "font_size": 10
    })

    blue_metric_format = workbook.add_format({
        "bg_color": "#D9E2F3",
        "border": 1,
        "font_size": 10
    })

    blue_value_format = workbook.add_format({
        "bg_color": "#D9E2F3",
        "border": 1,
        "align": "center",
        "font_size": 10
    })

    # --------------------------------------------------
    # COLUMN WIDTHS
    # --------------------------------------------------

    worksheet.set_column(0, 0, 45)

    worksheet.set_default_row(16)

    # --------------------------------------------------
    # STARTING ROW
    # --------------------------------------------------

    current_row = 0

    # --------------------------------------------------
    # PROCESS EACH FELLOWSHIP
    # --------------------------------------------------

    for program_name in sorted(program_files.keys()):

        data = program_files[program_name]

        display_name = data["display_name"]

        files = data["files"]

        # --------------------------------------------------
        # STORE METRICS BY DATE
        # --------------------------------------------------

        metrics_by_date = {}

        # --------------------------------------------------
        # PROCESS REPORTS
        # --------------------------------------------------

        for (
            file_date,
            file_path,
            file_name
        ) in files:

            try:

                df = pd.read_excel(
                    file_path,
                    sheet_name="Final"
                )

            except:
                continue

            progress_col = "Progress"

            if progress_col not in df.columns:
                continue

            # Normalize progress
            df[progress_col] = pd.to_numeric(
                df[progress_col],
                errors="coerce"
            ).fillna(0)

            # --------------------------------------------------
            # CALCULATE METRICS
            # --------------------------------------------------

            total = len(df)

            above = len(
                df[df[progress_col] >= 70]
            )

            below = len(
                df[df[progress_col] <= 69]
            )

            metrics_by_date[file_date] = {
                "total": total,
                "above": above,
                "below": below
            }

        # Skip empty programs
        if not metrics_by_date:
            continue

        # --------------------------------------------------
        # BUILD ARRAYS
        # --------------------------------------------------

        date_headers = []

        total_partials = []

        above_70 = []

        below_69 = []

        new_this_week = []

        converted_this_week = []

        # --------------------------------------------------
        # BUILD DATE-ALIGNED VALUES
        # --------------------------------------------------

        for file_date in all_dates:

            date_label = file_date.strftime(
                "%m/%d/%Y"
            )

            date_headers.append(date_label)

            # Existing week
            if file_date in metrics_by_date:

                total = metrics_by_date[file_date]["total"]

                above = metrics_by_date[file_date]["above"]

                below = metrics_by_date[file_date]["below"]

            else:

                total = 0
                above = 0
                below = 0

            total_partials.append(total)

            above_70.append(above)

            below_69.append(below)

            # --------------------------------------------------
            # IMPORT CONVERSION COUNT
            # --------------------------------------------------

            conversion_value = (
                conversions_lookup
                .get(program_name, {})
                .get(date_label, 0)
            )

            converted_this_week.append(
                conversion_value
            )

        # --------------------------------------------------
        # CALCULATE WEEK-OVER-WEEK GROWTH
        # --------------------------------------------------

        for i in range(len(total_partials)):

            current_total = total_partials[i]

            if i == len(total_partials) - 1:

                difference = 0

            else:

                previous_total = total_partials[i + 1]

                difference = max(
                    current_total - previous_total,
                    0
                )

            new_this_week.append(
                difference
            )

        # --------------------------------------------------
        # WRITE HEADER ROW
        # --------------------------------------------------

        worksheet.write(
            current_row,
            0,
            display_name,
            header_format
        )

        for col, date_value in enumerate(
            date_headers,
            start=1
        ):

            worksheet.write(
                current_row,
                col,
                date_value,
                date_format
            )

            worksheet.set_column(
                col,
                col,
                16
            )

        current_row += 1

        # --------------------------------------------------
        # METRIC ROWS
        # --------------------------------------------------

        metric_rows = [

            (
                "Total Partial Entries",
                total_partials,
                metric_format,
                value_format
            ),

            (
                "Above 70%",
                above_70,
                metric_format,
                value_format
            ),

            (
                "Below 69%",
                below_69,
                metric_format,
                value_format
            ),

            (
                "New Partial Entries this week",
                new_this_week,
                green_metric_format,
                green_value_format
            ),

            (
                "Partial Entries Converted to Apps (This Week)",
                converted_this_week,
                blue_metric_format,
                blue_value_format
            )
        ]

        # --------------------------------------------------
        # WRITE METRIC TABLE
        # --------------------------------------------------

        for (
            metric_name,
            values,
            metric_fmt,
            value_fmt
        ) in metric_rows:

            worksheet.write(
                current_row,
                0,
                metric_name,
                metric_fmt
            )

            for col, value in enumerate(
                values,
                start=1
            ):

                worksheet.write(
                    current_row,
                    col,
                    value,
                    value_fmt
                )

            current_row += 1

        # --------------------------------------------------
        # SPACE BETWEEN PROGRAMS
        # --------------------------------------------------

        current_row += 2

# --------------------------------------------------
# FINISHED
# --------------------------------------------------

print("\n========================================")
print("WEEK TO WEEK REPORT GENERATED")
print("========================================")
print(output_path)