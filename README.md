# Python utility to run Nerdgraph recursively and export dashboards data

- Script Functionality Includes
  - Retrieving the list of dashboards and their respective GUIDs
  - Make use of output sheet to export the dashboards json file

### Installation 
- Python is required as pre-requisite, the scripts are tested with latest Python version, here I am creating the viurtual environment and then installing the dependencies. 
    
      https://github.com/siddy15/nr-dashboard-export-json-pdf-png.git
      cd nr-entities-export-csv-json
      python3 -m venv env
      pip3 install requests, openpyxl
    
### Set env variable API_KEY and export data
        export API_KEY="YOUR NR USER KEY"
    
###  Export the list of dashboards and respective GUIDs
- Export the list of dashboards using GraphQL.

- Usage: 
```python3 dashboardList.py```

###  Export the dashboards json data
- Export all dashboards json using GraphQL.

- Usage: 
    `python3 dashboardExport.py`