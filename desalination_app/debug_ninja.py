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
    "date_to": f"{YEAR}-01-07",  # Just first week for comparison
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
print(f"\nParameters:")
print(f"  Lat: {LAT}, Lon: {LON}")
print(f"  Year: {YEAR}")
print(f"  Capacity: {CAPACITY}")
print(f"  Height: {HEIGHT}")
print(f"  Turbine: {TURBINE}")
print(f"  Dataset: {DATASET}")
print(f"  Raw: {RAW}")

# Get JSON data (what your app uses)
print("\n" + "=" * 60)
print("Fetching JSON data (what your app uses)...")
response = requests.get(url, params=params, headers=headers, timeout=60)
data = response.json()

raw_data = data.get("data", {})
print(f"Number of records: {len(raw_data)}")

# Simulate your app's export logic
rows = []
for timestamp, values in raw_data.items():
    if isinstance(values, dict):
        row = {"timestamp": timestamp}
        row.update(values)
        rows.append(row)

if rows:
    columns = list(rows[0].keys())

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

print("\n--- Your App Export (first 10 rows) ---")
print(f"{'timestamp':<20} {'electricity':<12} {'wind_speed':<12}")
for row in formatted_rows[:10]:
    print(f"{row['timestamp']:<20} {row['electricity']:<12} {row['wind_speed']:<12}")

# Get CSV directly from API (what the website uses)
print("\n" + "=" * 60)
print("Fetching CSV directly from API (official export)...")
params["format"] = "csv"
response_csv = requests.get(url, params=params, headers=headers, timeout=60)

# Parse CSV
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

print(f"Number of records: {len(csv_data)}")
print("\n--- Official API CSV (first 10 rows) ---")
print(f"{'time':<20} {'electricity':<12} {'wind_speed':<12}")
for row in csv_data[:10]:
    print(f"{row['time']:<20} {row['electricity']:<12} {row['wind_speed']:<12}")

# Compare values
print("\n" + "=" * 60)
print("COMPARISON (checking if values match):")
print("=" * 60)

all_match = True
for i, (app_row, api_row) in enumerate(zip(formatted_rows, csv_data)):
    # Compare electricity (allow small float difference)
    elec_match = abs(app_row['electricity'] - api_row['electricity']) < 0.001
    wind_match = abs(app_row['wind_speed'] - api_row['wind_speed']) < 0.001
    
    if not elec_match or not wind_match:
        all_match = False
        print(f"Row {i}: MISMATCH!")
        print(f"  App:   elec={app_row['electricity']}, wind={app_row['wind_speed']}")
        print(f"  API:   elec={api_row['electricity']}, wind={api_row['wind_speed']}")

if all_match:
    print("✓ All values match between your app and official API!")
else:
    print("✗ Some values don't match - see details above")
