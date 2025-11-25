import openpyxl
from utils import excel_date_to_string, format_hour


def analyze_schedule(file_path, target_name, rate, tax_rate):
    """Reads schedule XLSX, finds colored shifts, computes pay."""

    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
    except Exception as e:
        print(f"❌ Error loading workbook: {e}")
        return

    current_header_map = {}
    current_day_str = ""
    found_any_shift = False

    weekly_hours = 0
    weekly_gross = 0
    weekly_net = 0

    for row in sheet.iter_rows():
        first_cell_val = str(row[0].value).strip() if row[0].value else ""
        row_values = [str(c.value).strip() if c.value else "" for c in row]

        # 1. Detect header row
        if "10-11" in row_values:
            current_header_map = {}

            date_from_serial = excel_date_to_string(first_cell_val)
            current_day_str = date_from_serial if date_from_serial else row_values[-1]

            for cell in row:
                if cell.value and "-" in str(cell.value):
                    current_header_map[cell.column] = str(cell.value).strip()
            continue

        # 2. Detect target staff row
        if target_name.lower() in first_cell_val.lower():
            if not current_header_map:
                continue

            shifts_found = []

            for cell in row:
                if cell.column in current_header_map:

                    fill = cell.fill
                    is_colored = False

                    if fill.patternType == "solid":
                        rgb = fill.start_color.rgb
                        if rgb not in [None, "00000000", "FFFFFFFF"]:
                            is_colored = True

                    if is_colored:
                        shifts_found.append(current_header_map[cell.column])

            # 3. Handle shift
            if shifts_found:
                found_any_shift = True

                start_slot = shifts_found[0]
                end_slot = shifts_found[-1]
                start_hour = start_slot.split("-")[0]
                end_hour = end_slot.split("-")[1]

                human_start = format_hour(start_hour)
                human_end = format_hour(end_hour)

                print(f"Pizza Hut shift on {current_day_str} at {human_start} to {human_end}")

                # Pay calculations
                worked_hours = int(end_hour) - int(start_hour)
                gross = worked_hours * rate
                net = gross * (1 - tax_rate)

                weekly_hours += worked_hours
                weekly_gross += gross
                weekly_net += net

    if not found_any_shift:
        print(f"❌ No colored shifts found for {target_name}.")
        return

    print("\n📊 WEEKLY TOTALS")
    print(f"• Hours: {weekly_hours}")
    print(f"• Gross: ${weekly_gross:.2f}")
    print(f"• Net: ${weekly_net:.2f}")
    print("=" * 40)
