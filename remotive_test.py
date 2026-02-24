## Importing necessary libraries to work with APIs and Excel files
import json
import requests
import openpyxl
from openpyxl import Workbook

wb = Workbook()

ws = wb.active

ws.title = "Remote Jobs"

url = "https://remotive.com/api/remote-jobs"

response = requests.get(url)

json_data = response.text

data = json.loads(json_data)

print(data)

serialized_json = json.dumps(data, indent=4, separators=(", ", ": "))

print(serialized_json)

for item in data["jobs"]:
    print(item)
    print(item["title"])
    print(item["company_name"])
    print(item["category"])
    print(item["job_type"])
    print(item["publication_date"])
    print(item["url"])
    
    ws.append([item["title"], item["company_name"], item["category"], item["job_type"], item["publication_date"], item["url"]])

wb.save("week_4/test_api.xlsx")
