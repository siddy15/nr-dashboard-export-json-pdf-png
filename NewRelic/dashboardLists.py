import requests
from openpyxl import Workbook
import os
import json

endpoint = "https://api.newrelic.com/graphql"
api_key = os.environ.get('API_KEY')
output_excel = "dashboard_guids.xlsx"
query_file = "dashboardListQuery.graphql"

def load_query(file_path):
    try:
        with open(file_path, "r") as file:
            return file.read()
    except FileNotFoundError:
        print(f"Error: Query file '{file_path}' not found.")
        exit(1)

def fetch_all_dashboards(query):
    all_dashboards = []
    cursor = None

    while True:
        variables = {"cursor": cursor} if cursor else {}
        payload = {
            "query": query,
            "variables": variables
        }

        response = requests.post(
            endpoint,
            headers={"API-Key": api_key, "Content-Type": "application/json"},
            json=payload
        )

        if response.status_code != 200:
            print(f"Request failed (status {response.status_code})")
            print(response.text)
            break

        try:
            data = response.json()
            results = data["data"]["actor"]["entitySearch"]["results"]
            dashboards = results.get("entities", [])
            cursor = results.get("nextCursor")
            all_dashboards.extend(dashboards)
            print(f"Fetched {len(dashboards)} dashboards (Total so far: {len(all_dashboards)})")

            if not cursor:
                break  
        except Exception as e:
            print(f"Error processing response: {e}")
            break

    return all_dashboards

def save_to_excel(data):
    filtered_data = [entity for entity in data if '/' in (entity.get("name") or '')]

    if not filtered_data:
        print("No dashboard names contain '/'. Nothing saved.")
        return

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Dashboards"
    sheet.append(["GUID", "Name", "Account ID"])

    for entity in filtered_data:
        sheet.append([
            entity.get("guid"),
            entity.get("name"),
            entity.get("accountId")
        ])

    workbook.save(output_excel)
    print(f"Filtered dashboards saved to '{output_excel}' ({len(filtered_data)} entries).")

if __name__ == "__main__":
    print("========== Loading GraphQL query ==========")
    query = load_query(query_file)

    print("========== Fetching dashboards with pagination ==========")
    dashboards = fetch_all_dashboards(query)

    if dashboards:
        print(f"Total dashboards retrieved: {len(dashboards)}")
        save_to_excel(dashboards)
    else:
        print("No dashboard data retrieved.")
