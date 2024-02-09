import requests
import json
import pandas as pd

url = "https://testimex.efes.am/webservice/policy"
data = pd.read_json('C:/Users/AramMartirosyan/OneDrive - EFES ICJSC/Desktop/INECO/NEWS/New_Format.json',  orient='index')[0]
data = dict(data)

payload = json.dumps(data)

headers = {
  'Content-Type': 'application/json',
  'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjb250ZXh0Ijp7ImNsaWVudCI6eyJpZCI6IjI3NyIsIm5hbWUiOiJJa2h0c3lhbmRyIn0sImVudiI6IlBST0QifSwiaXNzIjoid3d3LmltZXguYW0iLCJpYXQiOjE3MDc0ODE4NzEsImV4cCI6MTcwNzY1NDY3MX0.aWBYxeI0NnQZqfc63flgEbm2TAABnGPv_t6be2h_7wM'
}

response = requests.request("POST", url, headers=headers, data=payload)
if response.status_code == 200:
  result = json.loads(response.text)
  print(result)
else:
  print(response.text)



