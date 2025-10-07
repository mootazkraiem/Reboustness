import csv
import os
import json

# --- Local CSV extraction ---

def extract_from_csv(folder_path, id_col='id', desc_col='description'):
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

def extract_pkg_descriptions(folder_path):
    import api  # Only available in ECU-TEST scripting environment
    results = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.pkg'):
            pkg_path = os.path.join(folder_path, filename)
            package = api.ObjectApi.OpenPackage(pkg_path)
            desc = package.GetDescription()
            id = package.GetId() if hasattr(package, 'GetId') else filename
            results.append({'id': id, 'description': desc})
    return results

def write_ecutest_to_csv(pkg_results, csv_path):
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['ID', 'Description'])
        for entry in pkg_results:
            writer.writerow([entry['id'], entry['description']])

if __name__ == "__main__":
    # CHANGE THIS TO YOUR LOCAL PATH
    folder_path = r'C:\_git_tm\ROOT_GEN6\xxx'
    csv_path = r'C:\Users\Mootaz\Desktop\reboustness\reboustness.csv'
    
    pkg_results = extract_pkg_descriptions(folder_path)
    write_ecutest_to_csv(pkg_results, csv_path)

    print(f"Extracted {len(pkg_results)} entries to {csv_path}")
