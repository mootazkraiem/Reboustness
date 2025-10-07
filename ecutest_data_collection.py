import requests
import csv
import os
import json
import xlsxwriter

# ECU-TEST API Data Collection Script

# --- Local CSV extraction ---

def extract_from_csv(folder_path, id_col='id', desc_col='description'):
    """
    Extract all IDs and descriptions from CSV files in a folder.
    Args:
        folder_path (str): Path to the folder containing CSV files.
        id_col (str): Column name for the ID.
        desc_col (str): Column name for the description.
    Returns:
        list of dict: [{'id': ..., 'description': ...}, ...]
    """
    results = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.csv'):
            with open(os.path.join(folder_path, filename), encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    results.append({'id': row.get(id_col, ''), 'description': row.get(desc_col, '')})
    return results

# --- Local JSON extraction ---

def extract_from_json(folder_path, id_key='id', desc_key='description'):
    """
    Extract all IDs and descriptions from JSON files in a folder.
    Args:
        folder_path (str): Path to the folder containing JSON files.
        id_key (str): Key name for the ID.
        desc_key (str): Key name for the description.
    Returns:
        list of dict: [{'id': ..., 'description': ...}, ...]
    """
    results = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.json'):
            with open(os.path.join(folder_path, filename), encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list):
                    for item in data:
                        results.append({'id': item.get(id_key, ''), 'description': item.get(desc_key, '')})
                elif isinstance(data, dict):
                    results.append({'id': data.get(id_key, ''), 'description': data.get(desc_key, '')})
    return results

# --- Local ECU-TEST API using ObjectApi ---

def extract_descriptions_with_objectapi(folder_path):
    """
    Extracts all descriptions using the local ECU-TEST ObjectApi.
    Args:
        folder_path (str): Path to the ECU-TEST package folder.
    Returns:
        list of dict: [{'id': ..., 'description': ...}, ...]
    """
    # This is a placeholder for the actual ECU-TEST scripting environment.
    # In a real ECU-TEST Python environment, you would do:
    # import api
    # package = api.ObjectApi.OpenPackage(folder_path)
    # desc = package.GetDescription()
    # You may need to loop over test cases/packages as needed.
    # Example:
    # results = []
    # for pkg in api.ObjectApi.GetPackages(folder_path):
    #     desc = pkg.GetDescription()
    #     results.append({'id': pkg.GetId(), 'description': desc})
    # return results
    pass

def extract_pkg_descriptions(folder_path):
    """
    Extract all IDs and descriptions from .pkg files using ECU-TEST ObjectApi.
    Args:
        folder_path (str): Path to the folder containing .pkg files.
    Returns:
        list of dict: [{'id': ..., 'description': ...}, ...]
    """
    import api  # Only available in ECU-TEST scripting environment
    results = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.pkg'):
            pkg_path = os.path.join(folder_path, filename)
            package = api.ObjectApi.OpenPackage(pkg_path)
            desc = package.GetDescription()
            # Try to get the ID, fallback to filename if not available
            id = package.GetId() if hasattr(package, 'GetId') else filename
            results.append({'id': id, 'description': desc})
    return results

def write_ecutest_to_excel(pkg_results, excel_path):
    """
    Write extracted ECU-TEST IDs and descriptions to an Excel file.
    Args:
        pkg_results (list of dict): Extracted package results.
        excel_path (str): Path to the Excel file to be created/updated.
    """
    workbook = xlsxwriter.Workbook(excel_path)
    worksheet = workbook.add_worksheet('ECU test description')
    # Write headers
    worksheet.write(0, 0, 'ID')
    worksheet.write(0, 1, 'Description')
    # Write data
    for idx, entry in enumerate(pkg_results, start=1):
        worksheet.write(idx, 0, entry['id'])
        worksheet.write(idx, 1, entry['description'])
    workbook.close()

def write_ecutest_to_csv(pkg_results, csv_path):
    """
    Write all IDs and descriptions to a CSV file (no external libraries needed).
    Args:
        pkg_results (list): List of dicts with 'id' and 'description'.
        csv_path (str): Path to the output CSV file.
    """
    import csv
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['ID', 'Description'])
        for entry in pkg_results:
            writer.writerow([entry['id'], entry['description']])

if __name__ == "__main__":
    # Example usage for CSV
    # csv_results = extract_from_csv(r'C:\path\to\your\csv_folder')
    # for entry in csv_results:
    #     print(f"ID: {entry['id']}, Description: {entry['description']}")

    # Example usage for JSON
    # json_results = extract_from_json(r'C:\path\to\your\json_folder')
    # for entry in json_results:
    #     print(f"ID: {entry['id']}, Description: {entry['description']}")

    # Example usage for ECU-TEST ObjectApi (run inside ECU-TEST Python environment)
    # folder_path = r'C:\_git_tm\ROOT_GEN6\xxx'  # <-- CHANGE THIS to your local ECU-TEST package path
    # results = extract_descriptions_with_objectapi(folder_path)
    # for entry in results:
    #     print(f"ID: {entry['id']}, Description: {entry['description']}")

    # Example usage for .pkg extraction and Excel writing (run inside ECU-TEST Python environment)
    folder_path = r'C:\_git_tm\ROOT_GEN6\xxx'  # <-- CHANGE THIS to your local ECU-TEST package path
    excel_path = r'C:\Users\Mootaz\Desktop\reboustness\reboustness.xlsx'  # <-- CHANGE if needed
    pkg_results = extract_pkg_descriptions(folder_path)
    write_ecutest_to_excel(pkg_results, excel_path)
    print(f"Extracted {len(pkg_results)} entries to {excel_path} (sheet: ECU test description)")

    # Example usage for .pkg extraction and CSV writing (run inside ECU-TEST Python environment)
    # folder_path = r'C:\_git_tm\ROOT_GEN6\xxx'  # <-- CHANGE THIS to your local ECU-TEST package path
    # csv_path = r'C:\Users\Mootaz\Desktop\reboustness\ecutest_descriptions.csv'  # <-- CHANGE if needed
    # pkg_results = extract_pkg_descriptions(folder_path)
    # write_ecutest_to_csv(pkg_results, csv_path)
    # print(f"Extracted {len(pkg_results)} entries to {csv_path}")
