import requests
import pandas as pd
import os

input_excel_file = "dashboard_guids.xlsx"
output_excel_file = "newrelic_dashboard_widgets.xlsx"

endpoint = "https://api.newrelic.com/graphql"
api_key = os.environ.get('API_KEY')

try:
    guid_df = pd.read_excel(input_excel_file)
    guid_df.columns = [col.lower() for col in guid_df.columns]  # Normalize column names
    if 'guid' not in guid_df.columns:
        raise ValueError("Input Excel file must contain a column named 'guid' (case-insensitive).")
    guids = guid_df['guid'].dropna().tolist()
except Exception as e:
    print(f"Failed to load GUIDs: {e}")
    exit()

# Function to build GraphQL query for a single GUID
def build_query(guid):
    return {
        "query": f"""
        {{
          actor {{
            entities(guids: "{guid}") {{
              ... on DashboardEntity {{
                guid
                name
                pages {{
                  widgets {{
                    title
                    id
                  }}
                }}
              }}
            }}
          }}
        }}
        """
    }

        
all_widgets = []

for guid in guids:
    print(f"Querying widgets for dashboard GUID: {guid}")
    response = requests.post(endpoint, 
                             json=build_query(guid), 
                             headers={"API-Key": api_key, "Content-Type": "application/json"}
                             )

    if response.status_code == 200:
        try:
            data = response.json()
            entities = data.get("data", {}).get("actor", {}).get("entities", [])

            for entity in entities:
                dashboard_guid = entity.get("guid")
                dashboard_name = entity.get("name")
                for page in entity.get("pages", []):
                    for widget in page.get("widgets", []):
                        all_widgets.append({
                            "dashboard_guid": dashboard_guid,
                            "dashboard_name": dashboard_name,
                            "widget_title": widget.get("title"),
                            "widget_id": widget.get("id")
                        })
        except Exception as e:
            print(f"Error processing data for GUID {guid}: {e}")
    else:
        print(f"Failed to query GUID {guid} (status code {response.status_code})")

if all_widgets:
    df = pd.DataFrame(all_widgets)
    df.to_excel(output_excel_file, index=False)
    print(f"{len(all_widgets)} widgets saved to '{output_excel_file}'")
else:
    print("No widget data found for any GUID.")
