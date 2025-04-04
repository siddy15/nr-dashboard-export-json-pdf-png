import requests
from openpyxl import Workbook
import os

endpoint = "https://api.newrelic.com/graphql"
api_key = os.environ.get('API_KEY')  
output_excel = "dashboard_guids.xlsx"
query_file = "dashboardListQuery.graphql"

def load_query(file_path):
    try:
        with open(file_path, "r") as file:
            return file.read()
    except FileNotFoundError:
        print(f"Error: The query file '{file_path}' does not exist.")
        exit(1)

def fetch_dashboard_data(query):
    response = requests.post(
        endpoint,
        json={"query": query},
        headers={"API-Key": api_key, "Content-Type": "application/json"},
    )
    if response.status_code == 200:
        data = response.json()
        try:
            return data["data"]["actor"]["entitySearch"]["results"]["entities"]
        except KeyError:
            print("Error: Unexpected data format.")
            return []
    else:
        print(f"Error: Request failed with status code {response.status_code}")
        return []

def save_to_excel(data):
    filtered_data = [entity for entity in data if '/' in (entity.get("name") or '')]

    if not filtered_data:
        print("No dashboard names contain '/'. No data saved.")
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
    print(f"Filtered data saved to '{output_excel}' with {len(filtered_data)} dashboards.")

if __name__ == "__main__":
    print("========== Loading GraphQL query ==========")
    query = load_query(query_file)
    
    print("========== Fetching dashboard data ==========")
    dashboards = fetch_dashboard_data(query)
    
    if dashboards:
        total_dashboards = len(dashboards)
        print(f"Total dashboards retrieved: {total_dashboards}")
        save_to_excel(dashboards)
    else:
        print("No dashboard data retrieved.")
