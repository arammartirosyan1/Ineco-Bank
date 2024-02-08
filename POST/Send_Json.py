import requests
import json
import pandas as pd

url = "https://testimex.efes.am/webservice/policy"
data = pd.read_json('C:/Users/aramm/OneDrive - EFES ICJSC/Desktop/INECO/NEWS/New_Format.json',  orient='index')[0]
data = dict(data)

payload = json.dumps(data)

headers = {
  'Content-Type': 'application/json',
  'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjb250ZXh0Ijp7ImNsaWVudCI6eyJpZCI6IjI3NyIsIm5hbWUiOiJJa2h0c3lhbmRyIn0sImVudiI6IlBST0QifSwiaXNzIjoid3d3LmltZXguYW0iLCJpYXQiOjE3MDczMDE2MjgsImV4cCI6MTcwNzQ3NDQyOH0.IdqAwyyS-yCgkti6moQ4w_VMUo16p62VH_WCvqciBc0'
}

response = requests.request("POST", url, headers=headers, data=payload)
result = json.loads(response.text)
print(result)

