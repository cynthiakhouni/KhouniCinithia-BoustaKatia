import requests
import json

url = 'https://www.renewables.ninja/api/data/wind'
params = {
    'lat': 56.0,
    'lon': -3.0,
    'date_from': '2024-01-01',
    'date_to': '2024-01-03',  # Just 3 days for testing
    'capacity': 1.0,
    'height': 100,
    'turbine': 'Vestas V80 2000',
    'dataset': 'merra2',
    'format': 'json',
    'raw': 'true'
}
headers = {'Authorization': 'Token 92fffc450a4b4f379d5499c0205d161a65a091a6'}

print("Fetching data from Renewable Ninja API...")
response = requests.get(url, params=params, headers=headers, timeout=60)
data = response.json()

print('\n--- Metadata ---')
print(json.dumps(data.get('metadata', {}), indent=2))

print('\n--- Data Structure ---')
raw_data = data.get('data', {})
print(f'Data type: {type(raw_data)}')
print(f'Number of entries: {len(raw_data)}')

print('\n--- First 5 entries (raw) ---')
import itertools
for i, (ts, vals) in enumerate(itertools.islice(raw_data.items(), 5)):
    print(f'{ts}: {vals}')

print('\n--- Extracted rows (as done in app) ---')
rows = []
for timestamp, values in raw_data.items():
    if isinstance(values, dict):
        row = {"timestamp": timestamp}
        row.update(values)
        rows.append(row)

if rows:
    columns = list(rows[0].keys())
    print(f'Columns: {columns}')
    for i, row in enumerate(rows[:5]):
        print(f'Row {i}: {row}')
