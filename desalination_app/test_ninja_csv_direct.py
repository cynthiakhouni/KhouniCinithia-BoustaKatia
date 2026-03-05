import requests

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
    'format': 'csv',  # Request CSV directly
    'raw': 'true'
}
headers = {'Authorization': 'Token 92fffc450a4b4f379d5499c0205d161a65a091a6'}

print("Fetching CSV directly from Renewable Ninja API...")
response = requests.get(url, params=params, headers=headers, timeout=120)

output_file = 'test_ninja_direct_from_api.csv'
with open(output_file, 'wb') as f:
    f.write(response.content)

print(f'Saved to: {output_file}')
print('\nFirst 15 lines:')
with open(output_file, 'r') as f:
    for i, line in enumerate(f):
        if i >= 15:
            break
        print(line.strip())
