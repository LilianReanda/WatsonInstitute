# --------------------------------------------------
# IMPORTS
# --------------------------------------------------

import pandas as pd
import os
import re
from datetime import datetime
import warnings

warnings.simplefilter("ignore", UserWarning)

# --------------------------------------------------
# CONFIGURATION
# --------------------------------------------------

reports_folder = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\reports"

conversions_folder = (
    r"C:\Users\Emanuel\PyCharmMiscProject"
    r"\WatsonInstitute"
    r"\Partial-Entries-Converted-to-Apps"
)

salesforce_folder = r"C:\Users\Emanuel\PyCharmMiscProject\WatsonInstitute\salesforce"

output_folder = os.path.join(reports_folder, "Week-to-Week-Comparison")
os.makedirs(output_folder, exist_ok=True)

today = datetime.today().strftime("%m-%d-%Y")

output_path = os.path.join(
    output_folder,
    f"{today} - Week to Week Comparison.xlsx"
)

launch_date = datetime.strptime("03-30-2026", "%m-%d-%Y")

# --------------------------------------------------
# NORMALIZATION
# --------------------------------------------------

def normalize_program_name(name):
    name = name.lower()

    name = name.replace("westernunion", "western union")
    name = name.replace("wellsfargo", "wells fargo")

    name = re.sub(r"f26", "", name, flags=re.IGNORECASE)
    name = re.sub(r"all r&a", "", name, flags=re.IGNORECASE)
    name = re.sub(r"weekly apps", "", name, flags=re.IGNORECASE)
    name = re.sub(r"weekly", "", name, flags=re.IGNORECASE)

    name = re.sub(r"\s+", " ", name).strip()

    return name

# --------------------------------------------------
# REPORT FILES
# --------------------------------------------------

pattern = re.compile(
    r"(\d{2}-\d{2}-\d{4}) - (.+?) - Partial Entries Report.*\.xlsx$",
    re.IGNORECASE
)

program_files = {}

for root, dirs, files_in_dir in os.walk(reports_folder):

    if "Week-to-Week-Comparison" in root:
        continue

    for file in files_in_dir:

        match = pattern.match(file)
        if not match:
            continue

        date_str, program_name = match.groups()

        normalized_program = normalize_program_name(program_name)
        display_name = program_name.strip()

        try:
            file_date = datetime.strptime(date_str, "%m-%d-%Y")
        except:
            continue

        if file_date < launch_date:
            continue

        full_path = os.path.join(root, file)

        program_files.setdefault(normalized_program, {
            "display_name": display_name,
            "files": []
        })

        program_files[normalized_program]["files"].append(
            (file_date, full_path, file)
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
    r"(\d{2}-\d{2}-\d{2})-(.+?)-.*\.xlsx$",
    re.IGNORECASE
)

for root, dirs, files in os.walk(conversions_folder):

    for file in files:

        match = conversion_pattern.match(file)
        if not match:
            continue

        date_str, raw_program = match.groups()

        normalized_program = normalize_program_name(raw_program)

        try:
            file_date = datetime.strptime(date_str, "%m-%d-%y")
            formatted_date = file_date.strftime("%m/%d/%Y")
        except:
            continue

        try:
            df_conversion = pd.read_excel(os.path.join(root, file))
            conversion_count = len(df_conversion.index)
        except:
            conversion_count = 0

        conversions_lookup.setdefault(normalized_program, {})
        conversions_lookup[normalized_program][formatted_date] = conversion_count

# --------------------------------------------------
# SALESFORCE
# --------------------------------------------------

salesforce_lookup = {}

for root, dirs, files in os.walk(salesforce_folder):

    for file in files:

        if not file.lower().endswith(".xlsx"):
            continue

        file_path = os.path.join(root, file)

        try:
            df_sf = pd.read_excel(file_path)
        except:
            continue

        if "Application Date Submitted" not in df_sf.columns:
            continue

        raw_name = file.split("-2026")[0]
        normalized_program = normalize_program_name(raw_name)

        df_sf["Application Date Submitted"] = pd.to_datetime(
            df_sf["Application Date Submitted"],
            errors="coerce"
        )

        df_sf = df_sf.dropna(subset=["Application Date Submitted"])

        for app_date in df_sf["Application Date Submitted"]:

            days_since_sunday = (app_date.weekday() + 1) % 7
            sunday = app_date - pd.Timedelta(days=days_since_sunday)

            report_date = sunday + pd.Timedelta(days=8)
            report_label = report_date.strftime("%m/%d/%Y")

            # --------------------------------------------------
            # EXCEPCIÓN FIX
            # --------------------------------------------------
            if report_label == "05/05/2026":
                report_label = "05/04/2026"

            salesforce_lookup.setdefault(normalized_program, {})
            salesforce_lookup[normalized_program][report_label] = (
                salesforce_lookup[normalized_program].get(report_label, 0) + 1
            )

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

        data = program_files[program_name]
        display_name = data["display_name"]
        files = data["files"]

        metrics_by_date = {}

        for file_date, file_path, _ in files:

            try:
                df = pd.read_excel(file_path, sheet_name="Final")
            except:
                continue

            if "Progress" not in df.columns:
                continue

            df["Progress"] = pd.to_numeric(df["Progress"], errors="coerce").fillna(0)

            metrics_by_date[file_date] = {
                "total": len(df),
                "above": len(df[df["Progress"] >= 70]),
                "below": len(df[df["Progress"] <= 69])
            }

        if not metrics_by_date:
            continue

        date_headers = []
        total_partials = []
        above_70 = []
        below_69 = []
        new_this_week = []
        converted_this_week = []
        salesforce_completed = []
        conversion_percent = []

        for file_date in all_dates:

            label = file_date.strftime("%m/%d/%Y")
            date_headers.append(label)

            data = metrics_by_date.get(file_date, {"total": 0, "above": 0, "below": 0})

            total_partials.append(data["total"])
            above_70.append(data["above"])
            below_69.append(data["below"])

            converted_this_week.append(
                conversions_lookup.get(program_name, {}).get(label, 0)
            )

            sf_val = salesforce_lookup.get(program_name, {}).get(label, 0)
            salesforce_completed.append(sf_val)

        for i in range(len(total_partials)):
            if i == len(total_partials) - 1:
                new_this_week.append(0)
            else:
                new_this_week.append(
                    max(total_partials[i] - total_partials[i + 1], 0)
                )

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
            ("Salesforce Completed Applications", salesforce_completed, blue_metric_format, blue_value_format),
            ("Total Partial Entries", total_partials, metric_format, value_format),
            ("Above 70%", above_70, metric_format, value_format),
            ("Below 69%", below_69, metric_format, value_format),
            ("New Partial Entries this week", new_this_week, green_metric_format, green_value_format),
            ("Partial Entries Converted to Apps (This Week)", converted_this_week, white_metric_format, white_value_format),
            ("Percentage of Partial entry conversion vs weekly apps", conversion_percent, white_metric_format, percent_format)
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