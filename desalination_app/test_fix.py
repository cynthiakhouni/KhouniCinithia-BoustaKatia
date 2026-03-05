from datetime import datetime

# Test timestamps from the API
test_timestamps = ['1704067200000', '1704070800000', '1704074400000']

print("OLD (local time):")
for ts in test_timestamps:
    dt = datetime.fromtimestamp(float(ts) / 1000.0)
    print(f"  {ts} -> {dt.strftime('%Y-%m-%d %H:%M:%S')}")

print("\nNEW (UTC):")
for ts in test_timestamps:
    dt = datetime.utcfromtimestamp(float(ts) / 1000.0)
    print(f"  {ts} -> {dt.strftime('%Y-%m-%d %H:%M:%S')}")

print("\nExpected (from official API CSV):")
print("  -> 2024-01-01 00:00:00")
print("  -> 2024-01-01 01:00:00")
print("  -> 2024-01-01 02:00:00")
