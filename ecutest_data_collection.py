import requests
import csv
import os
import json

# ECU-TEST API Data Collection Script

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
    # TODO: Change this endpoint to match your ECU-TEST API for listing test cases in a folder
    url = f"{ecu_api_url}/folders/{folder_id}/testcases"  # <-- CHANGE THIS if your API uses a different path
    # TODO: Set authentication as required by your ECU-TEST server
    response = requests.get(url, auth=ecu_auth)  # <-- CHANGE THIS if your API uses tokens or headers
    results = []
    if response.ok:
        data = response.json()
        # TODO: Adjust this if your API returns a different JSON structure
        for testcase in data.get('testcases', []):  # <-- CHANGE 'testcases' if your API uses a different key
            results.append({
                'id': testcase.get('id'),  # <-- CHANGE 'id' if your API uses a different field name
                'description': testcase.get('description', '')  # <-- CHANGE 'description' if your API uses a different field name
            })
    else:
        print(f"Failed to fetch test cases: {response.status_code}")
    return results

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

if __name__ == "__main__":
    # Example usage (fill in your details):
    folder_id = "XX"  # <-- CHANGE THIS to your folder name or ID
    ecu_api_url = "http://your-ecutest-server/api"  # <-- CHANGE THIS to your ECU-TEST API base URL
    ecu_auth = ("username", "password")  # <-- CHANGE THIS to your authentication method
    data = fetch_ecutest_ids_and_descriptions(folder_id, ecu_api_url, ecu_auth)
    for entry in data:
        print(f"ID: {entry['id']}, Description: {entry['description']}")
