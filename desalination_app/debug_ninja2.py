"""
Debug script to compare app export with official renewables.ninja API CSV
"""
import requests
import csv
from datetime import datetime, timezone

# Same parameters as your app
LAT = 56.0
LON = -3.0
YEAR = "2024"
CAPACITY = 1.0
HEIGHT = 100
TURBINE = "Vestas V80 2000"
DATASET = "merra2"
RAW = True

url = "https://www.renewables.ninja/api/data/wind"

# Build API parameters (same as your app)
params = {
    "lat": LAT,
    "lon": LON,
    "date_from": f"{YEAR}-01-01",
    "date_to": f"{YEAR}-01-07",
    "capacity": CAPACITY,
    "height": HEIGHT,
    "turbine": TURBINE,
    "dataset": DATASET,
    "format": "json",
    "raw": "true" if RAW else "false"
}
headers = {"Authorization": "Token 92fffc450a4b4f379d5499c0205d161a65a091a6"}

print("=" * 60)
print("DEBUG: Comparing App Export vs Official API")
print("=" * 60)

# Get JSON data (what your app uses)
response = requests.get(url, params=params, headers=headers, timeout=60)
data = response.json()
raw_data = data.get("data", {})

# Simulate your app's export logic
rows = []
for timestamp, values in raw_data.items():
    if isinstance(values, dict):
        row = {"timestamp": timestamp}
        row.update(values)
        rows.append(row)

# Format timestamps (your app's logic)
def format_timestamp(value):
    try:
        if isinstance(value, str):
            value = float(value)
        if isinstance(value, (int, float)) and value > 1000000000000:
            return datetime.fromtimestamp(value / 1000.0, tz=timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError, OSError, OverflowError):
        pass
    return value

formatted_rows = []
for row in rows:
    new_row = dict(row)
    if "timestamp" in new_row:
        new_row["timestamp"] = format_timestamp(new_row["timestamp"])
    formatted_rows.append(new_row)

# Get CSV directly from API (what the website uses)
params["format"] = "csv"
response_csv = requests.get(url, params=params, headers=headers, timeout=60)

lines = response_csv.text.strip().split('\n')
csv_data = []
for line in lines:
    if line.startswith('#') or line.startswith('time'):
        continue
    parts = line.split(',')
    if len(parts) >= 3:
        csv_data.append({
            'time': parts[0],
            'electricity': float(parts[1]),
            'wind_speed': float(parts[2])
        })

# Compare values
print("\nCOMPARISON:")
print("-" * 60)

all_match = True
for i, (app_row, api_row) in enumerate(zip(formatted_rows, csv_data)):
    elec_match = abs(app_row['electricity'] - api_row['electricity']) < 0.001
    wind_match = abs(app_row['wind_speed'] - api_row['wind_speed']) < 0.001
    
    if not elec_match or not wind_match:
        all_match = False
        print(f"Row {i}: MISMATCH!")
        print(f"  App:   elec={app_row['electricity']}, wind={app_row['wind_speed']}")
        print(f"  API:   elec={api_row['electricity']}, wind={api_row['wind_speed']}")

if all_match:
    print("[OK] All values match between your app and official API!")
else:
    print("[ERROR] Some values don't match - see details above")

# Also save both files for manual comparison
with open('app_export.csv', 'w', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=['timestamp', 'electricity', 'wind_speed'])
    writer.writeheader()
    writer.writerows(formatted_rows)

with open('api_export.csv', 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(['time', 'electricity', 'wind_speed'])
    for row in csv_data:
        writer.writerow([row['time'], row['electricity'], row['wind_speed']])

print("\nFiles saved:")
print("  - app_export.csv (your app's format)")
print("  - api_export.csv (official API format)")
