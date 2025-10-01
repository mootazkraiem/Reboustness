# Placeholder: Fetch descriptions from Codebeamer by IDs
def fetch_descriptions_from_codebeamer(ids, cb_api_url, cb_auth):
    """
    Fetch descriptions for a list of IDs from Codebeamer REST API.
    Args:
        ids (list): List of requirement IDs to fetch.
        cb_api_url (str): Base URL for Codebeamer API.
        cb_auth (tuple): (username, password) or token for authentication.
    Returns:
        dict: {id: description}
    """
    # Example implementation (requires requests):
    # import requests
    # result = {}
    # for req_id in ids:
    #     resp = requests.get(f"{cb_api_url}/item/{req_id}", auth=cb_auth)
    #     if resp.ok:
    #         data = resp.json()
    #         result[req_id] = data.get('description', '')
    #     else:
    #         result[req_id] = ''
    # return result
    pass

# Placeholder: Fetch descriptions from ECU-TEST by IDs
def fetch_ecutest_ids_and_descriptions(folder_id, ecu_api_url, ecu_auth):
    """
    Fetch all test case IDs and descriptions under a given folder in ECU-TEST.
    Args:
        folder_id (str): The folder name or ID (e.g., 'XX').
        ecu_api_url (str): Base URL for ECU-TEST API.
        ecu_auth (tuple): (username, password) or token for authentication.
    Returns:
        list of dict: [{'id': ..., 'description': ...}, ...]
    """
    import requests
    # [33mTODO: Change this endpoint to match your ECU-TEST API for listing test cases in a folder[0m
    url = f"{ecu_api_url}/folders/{folder_id}/testcases"  # <-- [33mCHANGE THIS if your API uses a different path[0m
    # [33mTODO: Set authentication as required by your ECU-TEST server[0m
    response = requests.get(url, auth=ecu_auth)  # <-- [33mCHANGE THIS if your API uses tokens or headers[0m
    results = []
    if response.ok:
        data = response.json()
        # [33mTODO: Adjust this if your API returns a different JSON structure[0m
        for testcase in data.get('testcases', []):  # <-- [33mCHANGE 'testcases' if your API uses a different key[0m
            results.append({
                'id': testcase.get('id'),  # <-- [33mCHANGE 'id' if your API uses a different field name[0m
                'description': testcase.get('description', '')  # <-- [33mCHANGE 'description' if your API uses a different field name[0m
            })
    else:
        print(f"Failed to fetch test cases: {response.status_code}")
    return results
import json
import os
import sys
import subprocess
import openpyxl
import xlsxwriter
import difflib
import math

def load_data_json(json_path):
    if os.path.exists(json_path):
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def save_data_json(json_path, data):
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def read_input():
    base_path = r'C:\Users\Mootaz\Desktop\reboustness'
    desc1_path = os.path.join(base_path, 'desc1.txt')
    desc2_path = os.path.join(base_path, 'desc2.txt')
    try:
        with open(desc1_path, 'r', encoding='utf-8') as f1:
            lines = f1.readlines()
            if not lines:
                raise ValueError("desc1.txt is empty.")
            req_id = lines[0].strip().split(":")[0]
            desc1 = ' '.join(line.strip() for line in lines[1:])
    except Exception as e:
        print(f"Error reading desc1.txt: {e}")
        return None, None, None
    try:
        with open(desc2_path, 'r', encoding='utf-8') as f2:
            desc2 = f2.read().strip()
    except Exception as e:
        print(f"Error reading desc2.txt: {e}")
        return None, None, None
    return req_id, desc1, desc2

def compare_descs(desc1, desc2):
    words1 = desc1.split()
    words2 = desc2.split()
    diff = list(difflib.ndiff(words1, words2))
    diff_count = sum(1 for d in diff if d.startswith("- "))
    ratio = diff_count / len(words1) if words1 else 1.0
    return diff, ratio

def diff_to_text(diff):
    result = []
    for d in diff:
        code = d[:2]
        word = d[2:]
        if code == '- ':
            result.append(f"{word}[difference]")
        elif code == '+ ':
            result.append(f"{word}[added]")
    return '\n'.join(result) if result else "No differences"

def open_file(filename):
    if sys.platform == "win32":
        os.startfile(filename)
    elif sys.platform == "darwin":
        subprocess.call(["open", filename])
    else:
        subprocess.call(["xdg-open", filename])

def estimate_row_height(text, col_width):
    if not text:
        return 30
    lines = math.ceil(len(text) / (col_width * 1.2)) if col_width > 0 else 1
    return max(lines * 15, 30)  # minimum height 30 points

def rich_string_from_diff(diff, highlight_type, workbook):
    """
    Build a list of (format, text) pairs for xlsxwriter write_rich_string,
    highlighting deleted words in desc1 (red), added words in desc2 (green),
    normal words in black.
    """
    normal_fmt = workbook.add_format({'font_color': 'black'})
    deleted_fmt = workbook.add_format({'font_color': 'red', 'bold': True})
    added_fmt = workbook.add_format({'font_color': 'green', 'bold': True})

    rich_text = []
    for d in diff:
        code = d[:2]
        word = d[2:]
        if highlight_type == 'desc1':
            # Highlight deleted words red in desc1 only
            if code == '- ':
                rich_text.extend([deleted_fmt, word + ' '])
            else:
                # '  ' or '+ ' normal in desc1
                rich_text.extend([normal_fmt, word + ' '])
        else:  # desc2
            # Highlight added words green in desc2 only
            if code == '+ ':
                rich_text.extend([added_fmt, word + ' '])
            else:
                # '  ' or '- ' normal in desc2
                rich_text.extend([normal_fmt, word + ' '])
    if not rich_text:
        rich_text = [normal_fmt, "No differences"]
    return rich_text

def main():
    base_path = r'C:\Users\Mootaz\Desktop\reboustness'
    filename = os.path.join(base_path, "reboustness.xlsx")
    json_path = os.path.join(base_path, "reboustness_data.json")
    req_id, desc1, desc2 = read_input()
    if not all([req_id, desc1, desc2]):
        print("Missing input data.")
        return

    diff, diff_ratio = compare_descs(desc1, desc2)
    desc3 = diff_to_text(diff)

    if diff_ratio == 0:
        status = "No changes"
    elif diff_ratio < 0.5:
        status = "Minor changes"
    else:
        status = "Major changes"

    # Load all data from JSON
    existing_data = load_data_json(json_path)
    # Check duplicates before appending
    if any(
        (row and ((row.get('id', '') if isinstance(row, dict) else (row[0] if len(row) > 0 else '')) == req_id))
        for row in existing_data
    ):
        print(f"ID '{req_id}' already exists. Entry rejected.")
        return
    # Append new entry as a dict with diff stored
    new_row = {
        'id': req_id,
        'desc1': desc1,
        'desc2': desc2,
        'diff': diff,  # store the diff list, not just text
        'desc3': desc3,
        'status': status,
        'ratio': diff_ratio * 100
    }
    # If upgrading from old list format, convert old rows to dicts
    for i, row in enumerate(existing_data):
        if isinstance(row, list):
            existing_data[i] = {
                'id': row[0],
                'desc1': row[1],
                'desc2': row[2],
                'diff': [],
                'desc3': row[3],
                'status': row[4],
                'ratio': row[5]
            }
    existing_data.append(new_row)
    # Save all data back to JSON
    save_data_json(json_path, existing_data)

    use_temp = os.path.exists(filename)
    temp_filename = filename + ".tmp"
    # Check if file is open by trying to rename it (Windows locks open files)
    if use_temp:
        try:
            os.rename(filename, filename)  # Try renaming to itself
        except OSError:
            print(f"Error: The file '{filename}' is open in Excel. Please close it and try again.")
            return
        out_filename = temp_filename
    else:
        out_filename = filename
    workbook = xlsxwriter.Workbook(out_filename)
    worksheet = workbook.add_worksheet("Data")

    # Formats
    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#4F81BD', 'color': 'white',
        'align': 'center', 'valign': 'vcenter'
    })
    center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    percent_format = workbook.add_format({'num_format': '0.00%', 'align': 'center', 'valign': 'vcenter'})

    headers = ["ID", "Codebeamer Desc", "ECU Test Desc (highlighted)", "Description Diff", "Status", "Change Ratio (%)"]
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header, header_format)

    for row_idx, row in enumerate(existing_data, start=1):
        # Use dict keys for all fields
        r_id = row.get('id', '')
        d1 = row.get('desc1', '')
        d2 = row.get('desc2', '')
        d3 = row.get('desc3', '')
        st = row.get('status', '')
        ratio_pct = row.get('ratio', 0)
        diff = row.get('diff', [])

        # Sanitize all values to avoid accidental formulas
        def sanitize(val):
            if isinstance(val, str) and val.startswith("="):
                return "'" + val  # Prepend single quote to avoid formula
            return val
        r_id, d1, d2, d3, st = map(sanitize, [r_id, d1, d2, d3, st])

        worksheet.write(row_idx, 0, r_id, center_format)

        # Use stored diff for highlighting for all rows
        if diff:
            rich_text1 = rich_string_from_diff(diff, 'desc1', workbook)
            if not rich_text1 or not any(isinstance(x, str) and x.strip() for x in rich_text1):
                rich_text1 = [workbook.add_format({'font_color': 'black'}), "No differences"]
            rich_text2 = rich_string_from_diff(diff, 'desc2', workbook)
            if not rich_text2 or not any(isinstance(x, str) and x.strip() for x in rich_text2):
                rich_text2 = [workbook.add_format({'font_color': 'black'}), "No differences"]
        else:
            # fallback for legacy rows
            rich_text1 = [workbook.add_format({'font_color': 'black'}), d1 or ""]
            rich_text2 = [workbook.add_format({'font_color': 'black'}), d2 or ""]

        worksheet.write_rich_string(row_idx, 1, *rich_text1, wrap_format)
        worksheet.write_rich_string(row_idx, 2, *rich_text2, wrap_format)

        # Description diff cell: blank if major changes
        if st == "Major changes":
            worksheet.write(row_idx, 3, "", wrap_format)
            # Fill background red
            red_fill = workbook.add_format({'bg_color': '#FFC7CE'})
            worksheet.set_row(row_idx, None, red_fill)
        else:
            worksheet.write(row_idx, 3, d3 or "", wrap_format)

        worksheet.write(row_idx, 4, st, center_format)
        worksheet.write(row_idx, 5, (ratio_pct or 0)/100, percent_format)

        # Adjust row height based on longest cell
        max_len = max(len(str(d1 or "")), len(str(d2 or "")), len(str(d3 or "")))
        row_height = estimate_row_height(' ' * max_len, 50)
        worksheet.set_row(row_idx, row_height)

    # Set column widths
    worksheet.set_column(0, 0, 15)
    worksheet.set_column(1, 2, 50)
    worksheet.set_column(3, 3, 30)
    worksheet.set_column(4, 4, 15)
    worksheet.set_column(5, 5, 18)

    # Create Summary worksheet
    summary = workbook.add_worksheet("Summary")

    # Prepare summary data
    statuses = ["No changes", "Minor changes", "Major changes"]
    counts = {s: 0 for s in statuses}
    seen_ids = set()
    for row in existing_data:
        if isinstance(row, dict):
            req_id = row.get('id', '')
            st = row.get('status', '')
        elif isinstance(row, list):
            req_id = row[0] if len(row) > 0 else ''
            st = row[4] if len(row) > 4 else ''
        else:
            continue
        if req_id not in seen_ids:
            if st in counts:
                counts[st] += 1
            seen_ids.add(req_id)

    # Write summary table
    summary.write(0, 0, "Status", header_format)
    summary.write(0, 1, "Count", header_format)

    for i, status in enumerate(statuses, start=1):
        summary.write(i, 0, status)
        summary.write(i, 1, counts[status])

    # Pie chart for status counts
    pie = workbook.add_chart({'type': 'pie'})
    pie.add_series({
        'name': 'Change Status Distribution',
        'categories': ['Summary', 1, 0, len(statuses), 0],
        'values':     ['Summary', 1, 1, len(statuses), 1],
    'data_labels': {'percentage': True},
    })
    pie.set_title({'name': 'Change Status Distribution'})
    summary.insert_chart('D2', pie, {'x_scale': 1.5, 'y_scale': 1.5})

    # Write Change Ratio table
    start_row = len(statuses) + 3
    summary.write(start_row, 0, "Description ID", header_format)
    summary.write(start_row, 1, "Change Ratio (%)", header_format)
    for idx, row in enumerate(existing_data, start=start_row + 1):
        if isinstance(row, dict):
            summary.write(idx, 0, row.get('id', ''))
            summary.write(idx, 1, row.get('ratio', 0), workbook.add_format({'num_format': '0.00', 'align': 'center'}))
        elif isinstance(row, list):
            summary.write(idx, 0, row[0] if len(row) > 0 else '')
            summary.write(idx, 1, (row[5] if len(row) > 5 else 0), workbook.add_format({'num_format': '0.00', 'align': 'center'}))

    # Bar chart for change ratios
    bar = workbook.add_chart({'type': 'bar'})
    bar.add_series({
        'name': 'Change Ratios',
        'categories': ['Summary', start_row + 1, 0, start_row + len(existing_data), 0],
        'values': ['Summary', start_row + 1, 1, start_row + len(existing_data), 1],
        'data_labels': {'value': True},
    })
    bar.set_title({'name': 'Change Ratios for Descriptions'})
    bar.set_x_axis({'name': 'Ratio (%)'})
    bar.set_y_axis({'reverse': True})
    summary.insert_chart('D20', bar, {'x_scale': 1.5, 'y_scale': 1.5})

    workbook.close()
    # If using temp file, replace the original after closing
    if use_temp:
        try:
            if os.path.exists(filename):
                os.remove(filename)
            os.rename(temp_filename, filename)
        except Exception as e:
            print(f"Error replacing the Excel file: {e}")
            print(f"The new file is saved as '{temp_filename}'.")

    print(f"✅ Excel file '{filename}' saved with full history and inline word highlighting.")

    open_file(filename)

if __name__ == "__main__":
    main()
