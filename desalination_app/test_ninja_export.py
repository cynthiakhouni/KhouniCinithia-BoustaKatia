import requests
import csv
from datetime import datetime

url = 'https://www.renewables.ninja/api/data/wind'
params = {
    'lat': 56.0,
    'lon': -3.0,
    'date_from': '2024-01-01',
    'date_to': '2024-01-03',
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

raw_data = data.get('data', {})

# Extract rows (as done in app)
rows = []
for timestamp, values in raw_data.items():
    if isinstance(values, dict):
        row = {"timestamp": timestamp}
        row.update(values)
        rows.append(row)

if rows:
    columns = list(rows[0].keys())
    print(f'Columns: {columns}')
    print(f'Number of rows: {len(rows)}')

# Helper function to convert timestamp (from your app)
def format_timestamp(value):
    try:
        if isinstance(value, str):
            value = float(value)
        if isinstance(value, (int, float)) and value > 1000000000000:
            return datetime.fromtimestamp(value / 1000.0).strftime("%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError, OSError, OverflowError):
        pass
    return value

# Prepare formatted rows (as done in app)
formatted_rows = []
has_timestamp = "timestamp" in columns
for row in rows:
    new_row = dict(row)
    if has_timestamp and "timestamp" in new_row:
        new_row["timestamp"] = format_timestamp(new_row["timestamp"])
    formatted_rows.append(new_row)

# Export to CSV (as done in app)
output_file = 'test_ninja_export.csv'
with open(output_file, "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=columns)
    writer.writeheader()
    writer.writerows(formatted_rows)

print(f'\nExported to: {output_file}')
print('\nFirst 10 rows of exported CSV:')
with open(output_file, 'r') as f:
    for i, line in enumerate(f):
        if i >= 11:
            break
        print(line.strip())
