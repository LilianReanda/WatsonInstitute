# --------------------------------------------------
# IMPORTS
# --------------------------------------------------

import pandas as pd
import re
from datetime import datetime
import warnings
from pathlib import Path

warnings.simplefilter("ignore", UserWarning)


# --------------------------------------------------
# CONFIGURATION
# --------------------------------------------------

home = Path.home()

base_folder = home / "PyCharmMiscProject" / "WatsonInstitute"

reports_folder = base_folder / "reports"

conversions_folder = base_folder / "Partial-Entries-Converted-to-Apps"

salesforce_folder = base_folder / "salesforce"

output_folder = reports_folder / "00-Week-to-Week-Comparison"

# CHANGE: parents=True so it does not fail if a parent folder is missing
output_folder.mkdir(parents=True, exist_ok=True)

today = datetime.today().strftime("%m-%d-%Y")

output_path = output_folder / f"{today} - Week to Week Comparison.xlsx"

launch_date = datetime.strptime("03-30-2026", "%m-%d-%Y")

# From this date onward:
# "New Partial Entries this week" comes from the Weekly tab
weekly_cutoff_date = datetime.strptime("05-18-2026", "%m-%d-%Y")


# --------------------------------------------------
# REMOVE PREVIOUS WEEK-TO-WEEK REPORTS
# --------------------------------------------------

for old_report in output_folder.glob("* - Week to Week Comparison.xlsx"):

    if old_report != output_path:

        try:
            old_report.unlink()

        except Exception as e:
            print(f"Could not delete old report: {old_report.name}")
            print(e)


# --------------------------------------------------
# NORMALIZATION
# --------------------------------------------------

def normalize_program_name(name):

    name = str(name).lower()

    # ----------------------------------------------
    # CHANGE: remove the export timestamp that
    # Salesforce adds to the file name, e.g.
    # "-2026-10-05-10-22-12"
    # This must happen BEFORE symbols are removed,
    # otherwise the digits survive in the key.
    # ----------------------------------------------
    name = re.sub(
        r"-?\d{4}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2}",
        "",
        name
    )

    # Standardize known program names
    name = name.replace("westernunion", "western union")
    name = name.replace("wellsfargo", "wells fargo")

    # Remove program cycle: S26 / S27 -> (nothing)
    # So "Truist S27" and "Truist" both become "truist"
    name = re.sub(r"s\d{2}", "", name, flags=re.IGNORECASE)

    # Remove known naming variations
    name = re.sub(r"f26", "", name, flags=re.IGNORECASE)
    name = re.sub(r"all r&a", "", name, flags=re.IGNORECASE)
    name = re.sub(r"weekly apps", "", name, flags=re.IGNORECASE)
    name = re.sub(r"weekly", "", name, flags=re.IGNORECASE)

    # Remove spaces, hyphens and punctuation
    name = re.sub(r"[^a-z0-9]", "", name)

    return name


# --------------------------------------------------
# REPORT FILES
# --------------------------------------------------

pattern = re.compile(
    r"(\d{2}-\d{2}-\d{4}) - "
    r"(.+?) - Partial Entries Report.*\.xlsx$",
    re.IGNORECASE
)

weekly_pattern = re.compile(
    r"(\d{2}-\d{2}-\d{4}) - "
    r"(.+?) - Partial Entries Report - Weekly.*\.xlsx$",
    re.IGNORECASE
)

program_files = {}

weekly_files_lookup = {}


for file_path in reports_folder.rglob("*.xlsx"):

    # Do not process the Week-to-Week output itself
    if "00-Week-to-Week-Comparison" in str(file_path):
        continue

    file = file_path.name

    # --------------------------------------------------
    # WEEKLY FILES
    # --------------------------------------------------

    weekly_match = weekly_pattern.match(file)

    if weekly_match:

        date_str, program_name = weekly_match.groups()

        normalized_program = normalize_program_name(program_name)

        try:
            file_date = datetime.strptime(date_str, "%m-%d-%Y")

        except Exception:
            continue

        weekly_files_lookup.setdefault(normalized_program, {})

        weekly_files_lookup[normalized_program][file_date] = file_path

    # --------------------------------------------------
    # REGULAR REPORT FILES
    # --------------------------------------------------

    match = pattern.match(file)

    if not match:
        continue

    date_str, program_name = match.groups()

    normalized_program = normalize_program_name(program_name)

    display_name = program_name.strip()

    try:
        file_date = datetime.strptime(date_str, "%m-%d-%Y")

    except Exception:
        continue

    if file_date < launch_date:
        continue

    program_files.setdefault(
        normalized_program,
        {
            "display_name": display_name,
            "files": []
        }
    )

    program_files[normalized_program]["files"].append(
        (file_date, file_path, file)
    )


# --------------------------------------------------
# MASTER DATES
# --------------------------------------------------

all_dates = set()

for data in program_files.values():

    for file_date, _, _ in data["files"]:

        all_dates.add(file_date)

all_dates = sorted(all_dates, reverse=True)


# --------------------------------------------------
# CONVERSIONS
# --------------------------------------------------

conversions_lookup = {}

conversion_pattern = re.compile(
    r"(\d{2}-\d{2}-\d{2})-"
    r"(.+?)-partials-converted-to-apps-this-week\.xlsx$",
    re.IGNORECASE
)


for file_path in conversions_folder.rglob("*.xlsx"):

    file = file_path.name

    match = conversion_pattern.match(file)

    if not match:
        continue

    date_str, raw_program = match.groups()

    normalized_program = normalize_program_name(raw_program)

    try:
        file_date = datetime.strptime(date_str, "%m-%d-%y")

        formatted_date = file_date.strftime("%m/%d/%Y")

    except Exception:
        continue

    # --------------------------------------------------
    # COUNT CONVERTED RECORDS
    # --------------------------------------------------

    try:
        df_conversion = pd.read_excel(file_path)

        conversion_count = len(df_conversion.index)

    except Exception:
        conversion_count = 0

    conversions_lookup.setdefault(normalized_program, {})

    conversions_lookup[normalized_program][formatted_date] = conversion_count


# --------------------------------------------------
# SALESFORCE
# --------------------------------------------------

salesforce_lookup = {}


for file_path in salesforce_folder.rglob("*.xlsx"):

    file = file_path.name

    try:
        df_sf = pd.read_excel(file_path)

    except Exception:
        print(f"[Salesforce] Could not read: {file}")
        continue

    if "Application Date Submitted" not in df_sf.columns:
        print(f"[Salesforce] Skipped (no 'Application Date Submitted'): {file}")
        continue

    # --------------------------------------------------
    # PROGRAM NAME FROM SALESFORCE FILE
    #
    # Uses the complete filename. The export timestamp
    # (e.g. -2026-10-05-10-22-12) is removed inside
    # normalize_program_name, so:
    #
    # Truist S27 All R&A Weekly apps-2026-10-05-10-22-12
    #   -> truist
    # Wells Fargo S27 All R&A Weekly apps-2026-10-05-10-22-31
    #   -> wellsfargo
    # --------------------------------------------------

    raw_name = Path(file).stem

    normalized_program = normalize_program_name(raw_name)

    # --------------------------------------------------
    # CLEAN APPLICATION DATES
    # --------------------------------------------------

    df_sf["Application Date Submitted"] = pd.to_datetime(
        df_sf["Application Date Submitted"],
        errors="coerce"
    )

    df_sf = df_sf.dropna(subset=["Application Date Submitted"])

    # --------------------------------------------------
    # GROUP SALESFORCE APPLICATIONS BY REPORTING WEEK
    # --------------------------------------------------

    for app_date in df_sf["Application Date Submitted"]:

        days_since_sunday = (app_date.weekday() + 1) % 7

        sunday = app_date - pd.Timedelta(days=days_since_sunday)

        report_date = sunday + pd.Timedelta(days=8)

        report_label = report_date.strftime("%m/%d/%Y")

        # --------------------------------------------------
        # SPECIAL FIX
        # --------------------------------------------------

        if report_label == "05/05/2026":
            report_label = "05/04/2026"

        salesforce_lookup.setdefault(normalized_program, {})

        salesforce_lookup[normalized_program][report_label] = (
            salesforce_lookup[normalized_program].get(report_label, 0) + 1
        )


# --------------------------------------------------
# DEBUG: KEY MATCHING
# --------------------------------------------------

print("\nReport keys:     ", sorted(program_files.keys()))
print("Salesforce keys: ", sorted(salesforce_lookup.keys()))

unmatched = set(salesforce_lookup.keys()) - set(program_files.keys())

if unmatched:
    print("\n[WARNING] Salesforce keys with no matching report program:")
    for key in sorted(unmatched):
        print("  -", key)


# --------------------------------------------------
# EXCEL OUTPUT
# --------------------------------------------------

with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:

    workbook = writer.book

    worksheet = workbook.add_worksheet("Week to Week Comparison")

    writer.sheets["Week to Week Comparison"] = worksheet

    # --------------------------------------------------
    # FORMATS
    # --------------------------------------------------

    light_orange = "#FCE5CD"

    header_format = workbook.add_format({
        "bold": True,
        "font_color": "white",
        "bg_color": "#0B3A6E",
        "align": "center",
        "border": 1
    })

    date_format = workbook.add_format({
        "bold": True,
        "font_color": "white",
        "bg_color": "#0B3A6E",
        "align": "center",
        "border": 1
    })

    metric_format = workbook.add_format({
        "bg_color": light_orange,
        "border": 1
    })

    value_format = workbook.add_format({
        "bg_color": light_orange,
        "border": 1,
        "align": "center"
    })

    green_metric_format = workbook.add_format({
        "bg_color": "#D9EAD3",
        "border": 1
    })

    green_value_format = workbook.add_format({
        "bg_color": "#D9EAD3",
        "border": 1,
        "align": "center"
    })

    blue_metric_format = workbook.add_format({
        "bg_color": "#D9E2F3",
        "border": 1
    })

    blue_value_format = workbook.add_format({
        "bg_color": "#D9E2F3",
        "border": 1,
        "align": "center"
    })

    white_metric_format = workbook.add_format({
        "bg_color": "#FFFFFF",
        "border": 1
    })

    white_value_format = workbook.add_format({
        "bg_color": "#FFFFFF",
        "border": 1,
        "align": "center"
    })

    percent_format = workbook.add_format({
        "num_format": "0.0%",
        "bg_color": "#FFFFFF",
        "border": 1,
        "align": "center"
    })

    worksheet.set_column(0, 0, 45)

    worksheet.set_default_row(16)

    current_row = 0

    # --------------------------------------------------
    # MAIN LOOP
    # --------------------------------------------------

    for program_name in sorted(program_files.keys()):

        program_data = program_files[program_name]

        display_name = program_data["display_name"]

        files = program_data["files"]

        # --------------------------------------------------
        # PARTIAL ENTRY METRICS
        # --------------------------------------------------

        metrics_by_date = {}

        for file_date, file_path, _ in files:

            try:
                df = pd.read_excel(file_path, sheet_name="Final")

            except Exception:
                continue

            if "Progress" not in df.columns:
                continue

            df["Progress"] = (
                pd.to_numeric(df["Progress"], errors="coerce")
                .fillna(0)
            )

            metrics_by_date[file_date] = {
                "total": len(df),
                "above": len(df[df["Progress"] >= 70]),
                "below": len(df[df["Progress"] <= 69])
            }

        if not metrics_by_date:
            continue

        # --------------------------------------------------
        # METRIC ARRAYS
        # --------------------------------------------------

        date_headers = []
        total_partials = []
        above_70 = []
        below_69 = []
        new_this_week = []
        converted_this_week = []
        salesforce_completed = []
        conversion_percent = []

        # --------------------------------------------------
        # BUILD METRICS BY DATE
        # --------------------------------------------------

        for file_date in all_dates:

            label = file_date.strftime("%m/%d/%Y")

            date_headers.append(label)

            date_metrics = metrics_by_date.get(
                file_date,
                {"total": 0, "above": 0, "below": 0}
            )

            total_partials.append(date_metrics["total"])
            above_70.append(date_metrics["above"])
            below_69.append(date_metrics["below"])

            # --------------------------------------------------
            # CONVERSIONS
            # --------------------------------------------------

            converted_this_week.append(
                conversions_lookup.get(program_name, {}).get(label, 0)
            )

            # --------------------------------------------------
            # SALESFORCE COMPLETED
            # --------------------------------------------------

            salesforce_completed.append(
                salesforce_lookup.get(program_name, {}).get(label, 0)
            )

        # --------------------------------------------------
        # NEW PARTIALS LOGIC
        # --------------------------------------------------

        for i, file_date in enumerate(all_dates):

            # ----------------------------------------------
            # NEW LOGIC USING WEEKLY TAB
            # ----------------------------------------------

            if file_date >= weekly_cutoff_date:

                weekly_file = (
                    weekly_files_lookup
                    .get(program_name, {})
                    .get(file_date)
                )

                if weekly_file:

                    try:
                        df_weekly = pd.read_excel(
                            weekly_file,
                            sheet_name="Weekly"
                        )

                        weekly_count = len(df_weekly.index)

                    except Exception:
                        weekly_count = 0

                else:
                    weekly_count = 0

                new_this_week.append(weekly_count)

            # ----------------------------------------------
            # OLD LOGIC
            # ----------------------------------------------

            else:

                if i == len(total_partials) - 1:

                    new_this_week.append(0)

                else:

                    new_this_week.append(
                        max(total_partials[i] - total_partials[i + 1], 0)
                    )

        # --------------------------------------------------
        # CONVERSION %
        # --------------------------------------------------

        for i in range(len(converted_this_week)):

            sf = salesforce_completed[i]

            conv = converted_this_week[i]

            conversion_percent.append((conv / sf) if sf else 0)

        # --------------------------------------------------
        # HEADER
        # --------------------------------------------------

        worksheet.write(current_row, 0, display_name, header_format)

        for col, date_value in enumerate(date_headers, start=1):

            worksheet.write(current_row, col, date_value, date_format)

            worksheet.set_column(col, col, 20)

        current_row += 1

        # --------------------------------------------------
        # METRICS
        # --------------------------------------------------

        metric_rows = [
            (
                "Salesforce Completed Applications",
                salesforce_completed,
                blue_metric_format,
                blue_value_format
            ),
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
                white_metric_format,
                white_value_format
            ),
            (
                "Percentage of Partial entry conversion vs weekly apps",
                conversion_percent,
                white_metric_format,
                percent_format
            )
        ]

        for metric_name, values, m_fmt, v_fmt in metric_rows:

            worksheet.write(current_row, 0, metric_name, m_fmt)

            for col, value in enumerate(values, start=1):

                worksheet.write(current_row, col, value, v_fmt)

            current_row += 1

        current_row += 2


# --------------------------------------------------
# DONE
# --------------------------------------------------

print("\n========================================")
print("WEEK TO WEEK REPORT GENERATED")
print("========================================")
print(output_path)