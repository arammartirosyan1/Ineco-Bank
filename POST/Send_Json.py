import requests
import json
import pandas as pd

url = "https://testimex.efes.am/webservice/policy"
data = pd.read_json('C:/Users/AramMartirosyan/OneDrive - EFES ICJSC/Desktop/INECO/ACCIDENT/News/PA_Format0.json',  orient='index')[0]
data = dict(data)

payload = json.dumps(data)

headers = {
  'Content-Type': 'application/json',
  'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjb250ZXh0Ijp7ImNsaWVudCI6eyJpZCI6IjI3NyIsIm5hbWUiOiJJa2h0c3lhbmRyIn0sImVudiI6IlBST0QifSwiaXNzIjoid3d3LmltZXguYW0iLCJpYXQiOjE3MDc3NDUzOTYsImV4cCI6MTcwNzkxODE5Nn0.KnEVjdf3m8BC1at-3_qbmWKODXu88JovFaJ8ohmeWJQ'
}

response = requests.request("POST", url, headers=headers, data=payload)
if response.status_code == 200:
    print("Success")
    result = json.loads(response.text)
    print(result)
else:
    print("Error")
    print(response.text, response.status_code)



