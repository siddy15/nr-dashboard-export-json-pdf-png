import requests
from openpyxl import Workbook
import os

# Configurations
endpoint = "https://api.newrelic.com/graphql" 

# Set the API_key as environment variable
api_key = os.environ.get('API_KEY')

# Name of the output file
output_excel = "dashboard_guids.xlsx"  

query_file = "dashboardListQuery.graphql"  

# Loading the GraphQL
def load_query(file_path):
    try:
        with open(file_path, "r") as file:
            return file.read()
    except FileNotFoundError:
        print(f"Error: The query file '{file_path}' does not exist.")
        exit(1)

# Fetching the data
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

# Saving the data to external file 
def save_to_excel(data):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Dashboards"
    
    sheet.append(["GUID", "Name", "Account ID"])
    
    for entity in data:
        sheet.append([entity.get("guid"), entity.get("name"), entity.get("accountId")])
    
    workbook.save(output_excel)
    print(f"Data saved successfully to {output_excel}")

if __name__ == "__main__":
    print("========== Loading GraphQL query ==========")
    query = load_query(query_file)
    
    print("========== Fetching dashboard data ==========")
    dashboards = fetch_dashboard_data(query)
    if dashboards:
        print(f"Retrieved {len(dashboards)} dashboards.")
        save_to_excel(dashboards)
    else:
        print("========== No dashboard data retrieved ==========")
