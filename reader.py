import csv
import json
import string
import openpyxl
from dateutil import parser
from datetime import date, datetime
from tkinter import filedialog, Tk
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# var
root = Tk()
root.withdraw()

valid_statuses = {"completed", "incomplete", "overdue"}
required_headers = {
    "employee_id",
    "name",
    "location",
    "training_course",
    "assigned_date",
    "due_date",
    "completion_date",
    "status"
}

valid_warning = {"assigned_date", "due_date", "status"}
valid_errors = {"employee_id", "name", "location", "training_course"}

count_completed = 0
count_uncompleted = 0
count_overdue = 0
count_invalid = 0

file_path = filedialog.askopenfilename(
    title="Select CSV File",
    filetypes=[("CSV File", "*.csv")]
)

def is_true_csv(test_file_path):
    try:
        with open(test_file_path, mode='r', encoding='utf-8', errors='strict') as file:

            # test 1
            test_sample = file.read(4096)

            # If the file contains non-printable binary characters, it is not a true text CSV.
            if not all(c in string.printable or c.isprintable() for c in test_sample):
                return False, None, "File contains invalid binary characters (not a text file)."

            # test 2: Checking with Sniffer
            dialect = csv.Sniffer().sniff(test_sample)
            file.seek(0)
            reader = csv.reader(file, dialect)

            # get the column count of the first line
            first_row = next(reader)
            expected_columns = len(first_row)
            if expected_columns <= 1:
                return False, None, "Could not find reliable tabular columns (only found 1 or 0 columns)."

            # Verify the next five rows.
            for i, row in enumerate(reader):
                if i > 4:  # Checking 5 rows is usually enough to verify standard consistency
                    break
                if len(row) != expected_columns:
                    return False, None, f"Row formatting is inconsistent. Line {i + 2} has a different column count."

            return True, dialect, f"Valid CSV. Detected delimiter: '{dialect.delimiter}'"

    except (csv.Error, TypeError, ValueError):
        return False, None, "Failed CSV structural validation (corrupted or malformed syntax)."
    except UnicodeDecodeError:
        return False, None, "Failed text encoding. File is likely a binary file or not UTF-8."
    except FileNotFoundError:
        return False, None, "The specified file does not exist."


def check_headers(headers):

    # Clean headers to lowercase
    clean_header = {h.strip().lower() for h in headers}

    # check missing headers
    missing_header = required_headers - clean_header

    if missing_header:
        return False, f"Missing Required columns: {', '.join(missing_header)}"

    return True, "All required columns are present"


def check_overdue(item):
    warning = ""
    assigned_date_str = str(item.get("assigned_date", "")).strip()
    due_date_str = str(item.get("due_date", "")).strip()

    completion_date_str = str(item.get("completion_date", "")).strip()
    if not completion_date_str:
        completion_date_str = "Not set"

    status = str(item.get("status", "")).strip().lower()
    if status not in valid_statuses:
        warning = f"Invalid status: {status}"
        status = "invalid"

    days_overdue = 0

    today = date.today()

    if due_date_str and status != "completed":
        try:
            dd = parser.parse(due_date_str).date()

            if dd < today:
                days_overdue = (today - dd).days

                if status == "incomplete":
                    status = "overdue"


        except (ValueError, TypeError):
            print(f"Warning: Could not parse date format")


    return {
        "completion_date": completion_date_str,
        "assigned_date": assigned_date_str,
        "due_date": due_date_str,

        "status": status,
        "days_overdue": days_overdue,
        "warning": warning
    }


def table_validation(table):
    result_table = {}

    for item in table:

        # If Location is not Valid.
        if str(item['location']).strip():
            location = item.get("location")
        else:
            location = "Unknown location"

        # If Name is not Valid.
        if str(item['name']).strip():
            employee_name = item.get("name")
        else:
            employee_name = "Unknown Employee"

        # If ID is not Valid.
        if str(item['employee_id']).strip():
            employee_id = item.get("employee_id")
        else:
            employee_id = f"Error: {employee_name}"

        if location not in result_table:
            # Create Table
            result_table[location] = {}

        if employee_id not in result_table[location]:
            result_table[location][employee_id] = {}
            result_table[location][employee_id][employee_name] = {}

            result_table[location][employee_id][employee_name]["Error"] = []
            result_table[location][employee_id][employee_name]["Warnings"] = []
            result_table[location][employee_id][employee_name]["Training"] = {}

        # training over-due?
        training_course = str(item.get("training_course", "") or "").strip()

        if not training_course:
            training_course = "Unknown Training"


        result = check_overdue(item)

        if result["warning"]:
            result_table[location][employee_id][employee_name]["Warnings"].append(result["warning"])

        if training_course in result_table[location][employee_id][employee_name]["Training"]:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            training_course = training_course + "_" + timestamp

        result_table[location][employee_id][employee_name]["Training"][training_course] = {
            "Completion_Date": result["completion_date"],
            "Assigned_Date": result["assigned_date"],
            "Due_Date": result["due_date"],

            "Status": result["status"],
            "Days_Over_Due": result["days_overdue"]
        }


        # check for missing
        for key, value in item.items():
            if key == "completion_date":
                status = str(item.get("status", "")).strip().lower()

                if status == "completed" and not str(value).strip():
                    result_table[location][employee_id][employee_name]["Warnings"].append(
                        "Completed training has no completion date"
                    )

                continue

            if not str(value).strip():

                if key in valid_warning:
                    result_table[location][employee_id][employee_name]["Warnings"].append(key)

                elif key in valid_errors:
                    result_table[location][employee_id][employee_name]["Error"].append(key)


    return result_table


def save_file(r_table):
    save_path = filedialog.asksaveasfilename(
        title="Save Results",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )

    if save_path:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Training Report"

        # Define Styles
        # Font sizes and colors
        header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        data_font = Font(name="Arial", size=11)
        alert_font = Font(name="Arial", size=11, color="9C0006")  # Dark red text

        # Background fills
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")  # Dark Blue
        alert_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Light Red

        # Alignments
        center_align = Alignment(horizontal="center", vertical="center")
        left_align = Alignment(horizontal="left", vertical="center")

        # Borders (thin gray lines)
        thin_side = Side(border_style="thin", color="D3D3D3")
        cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        Headers = ["location", "employee_id", "name", "training_course", "status", "assigned_date", "due_date", "completion_date", "Days_Over_Due"]
        ws.append(Headers)

        # Format the header row (Row 1)
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = cell_border


        # 4. Flatten and iterate through your nested result_table object
        for location, id_dict in r_table.items():
            for emp_id, name_dict in id_dict.items():
                for emp_name, categories in name_dict.items():
                    training_data = categories.get("Training", {})

                    for item, value in training_data.items():

                        # Append data row
                        ws.append([
                            location,
                            emp_id,
                            emp_name,
                            item,
                            value["Status"],
                            value["Assigned_Date"],
                            value["Due_Date"],
                            value["Completion_Date"],
                            value["Days_Over_Due"],
                        ])

                        # Get the row index just appended
                        current_row = ws.max_row

                        # Apply standard styles to the new row
                        ws[f"A{current_row}"].alignment = center_align  # location
                        ws[f"B{current_row}"].alignment = left_align  # employee_id
                        ws[f"C{current_row}"].alignment = center_align  # name
                        ws[f"D{current_row}"].alignment = center_align  # training_course
                        ws[f"E{current_row}"].alignment = center_align  # status
                        ws[f"F{current_row}"].alignment = center_align  # assigned_date
                        ws[f"G{current_row}"].alignment = center_align  # due_date
                        ws[f"H{current_row}"].alignment = center_align  # completion_date
                        ws[f"I{current_row}"].alignment = center_align  # Days Over Due

                        for col in ["A", "B", "C", "D", "E", "F", "G", "H", "I"]:
                            cell = ws[f"{col}{current_row}"]
                            cell.font = data_font
                            cell.border = cell_border

                        # Conditional Formatting: Highlight in Red if they are Overdue
                        if value["Days_Over_Due"] >= 1:
                            ws[f"C{current_row}"].fill = alert_fill
                            ws[f"C{current_row}"].font = alert_font
                            ws[f"D{current_row}"].fill = alert_fill
                            ws[f"D{current_row}"].font = alert_font
                            ws[f"I{current_row}"].fill = alert_fill
                            ws[f"I{current_row}"].font = alert_font

        # 5. Automatically adjust column widths so data isn't cut off
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        wb.save(save_path)

        print("✅ File saved successfully!")
    else:
        print("❌ Save operation cancelled by user.")



# ======================================= Read File ================================================================

if not file_path:
    print("No file was selected.")
    raise SystemExit

is_valid, dialect, message = is_true_csv(file_path)


if file_path and is_valid:
    print(f"Selected File: {file_path}")

    with open(file_path, mode='r', encoding='utf-8') as file:
        sample = file.read(2048)  # set sample size
        file.seek(0)  # move cursor to 0 just in incase

        # create a CSV Reader object
        csv_reader = csv.reader(file, dialect)

        # check if has header
        has_header = csv.Sniffer().has_header(sample)

        if has_header:
            header = next(csv_reader)

            # check headers
            check, message = check_headers(header)

            if check:
                print(message)

                data_rows = []

                file.seek(0)
                dict_reader = csv.DictReader(file, dialect=dialect)

                # read the status row and count the status
                for row in dict_reader:
                    # Guard against empty or malformed rows
                    if not row:
                        continue

                    normalized_row = {
                        str(key).strip().lower(): value
                        for key, value in row.items()
                        if key is not None
                    }

                    data_rows.append(normalized_row)


                validation_results = table_validation(data_rows)

                for location, id_dict in validation_results.items():
                    for emp_id, name_dict in id_dict.items():
                        for emp_name, categories in name_dict.items():
                            training_data = categories.get("Training", {})

                            for item, value in training_data.items():
                                status = str(value.get("Status", "")).strip().lower()

                                if status == "completed":
                                    count_completed += 1

                                elif status == "incomplete":
                                    count_uncompleted += 1

                                elif status == "overdue":
                                    count_overdue += 1

                                elif status == "invalid":
                                    count_invalid += 1


                total_rows = len(data_rows)
                percentage = count_completed / total_rows if total_rows else 0

                print(f"\nTotal Rows: {total_rows}")
                print("========== Counts ==========")
                print(f"\nCompleted: {count_completed}")
                print(f"Incomplete: {count_uncompleted}")
                print(f"Overdue: {count_overdue}\n")
                print(f"Invalid: {count_invalid}")

                print(f"% completed: {percentage:.1%}\n")
                print("====== Missing Items ======")
                print(json.dumps(validation_results, indent=4))

                print("============================\n")

                save_file(validation_results)

            elif not check:
                print(message)

        else:
            print("No Header")

            row_count = sum(1 for row in csv_reader)
            print(f"Total Rows: {row_count}")
            file.seek(0)  # move cursor to 0 before reading again

            for row in csv_reader:
                print(row)

else:
    if message:
        print(message)


