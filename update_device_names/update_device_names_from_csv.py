import requests
import csv
import json

API_KEY = 'YOUR_API_KEY'  # Replace with your actual API key
API_URL = 'https://api-us-east-1-cell-1.domotz.com/public-api/v1'
INPUT_CSV = 'device_name_discrepancies.csv'
OUTPUT_CSV = 'updated_devices.csv'

def update_device_name(agent_id, device_id, new_name):
    url = f'{API_URL}/agent/{agent_id}/device/{device_id}/user_data/name'
    headers = {
        'x-api-key': API_KEY,
        'Content-Type': 'application/json'
    }
    data = new_name  # Ensure new_name is a string
    response = requests.put(url, headers=headers, data=json.dumps(data))
    
    # Log request and response details
    print(f"Request URL: {url}")
    print(f"Request Headers: {headers}")
    print(f"Request Data: {json.dumps(data)}")
    print(f"Response Status Code: {response.status_code}")
    print(f"Response Text: {response.text}")
    
    if response.status_code == 200 or response.status_code == 204:
        return True
    elif response.status_code == 423:
        print(f"Locked: {response.status_code} - {response.text}")
    else:
        print(f"Failed to update device: {response.status_code} - {response.text}")
    
    return False

def read_csv(filename):
    devices = []
    with open(filename, mode='r') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            devices.append(row)
    return devices

def save_to_csv(devices, filename='updated_devices.csv'):
    with open(filename, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Agent ID', 'Agent Name', 'Device ID', 'Old Device Name', 'New DHCP Name', 'Status'])
        for device in devices:
            writer.writerow([device['Agent ID'], device['Agent Name'], device['Device ID'], device['Device Name'], device['DHCP Name'], device['Status']])

def main():
    devices_to_update = read_csv(INPUT_CSV)
    updated_devices = []
    for device in devices_to_update:
        try:
            if update_device_name(device['Agent ID'], device['Device ID'], device['DHCP Name']):
                device['Status'] = 'Updated'
            else:
                device['Status'] = 'Failed: Could not update device name'
        except requests.HTTPError as e:
            device['Status'] = f'Failed: {str(e)}'
        updated_devices.append(device)
    save_to_csv(updated_devices, OUTPUT_CSV)

if __name__ == '__main__':
    main()
