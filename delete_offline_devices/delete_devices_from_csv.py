import requests
import csv

API_KEY = 'YOUR_API_KEY'  # Replace with your actual API key
API_URL = 'https://api-us-east-1-cell-1.domotz.com/public-api/v1'
INPUT_CSV = 'offline_devices.csv'
OUTPUT_CSV = 'deleted_devices.csv'

def delete_device(agent_id, device_id):
    url = f'{API_URL}/agent/{agent_id}/device/{device_id}'
    headers = {'x-api-key': API_KEY}
    response = requests.delete(url, headers=headers)
    
    # Check if the response status code indicates success
    if response.status_code == 204:
        return True
    elif response.status_code == 403:
        # Handle specific status codes separately if needed
        print(f"Forbidden: {response.status_code} - {response.text}")
    else:
        print(f"Failed to delete device: {response.status_code} - {response.text}")
    
    return False

def read_csv(filename):
    devices = []
    with open(filename, mode='r') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            devices.append(row)
    return devices

def save_to_csv(devices, filename='deleted_devices.csv'):
    with open(filename, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Agent ID', 'Agent Name', 'Device ID', 'Device Name', 'IP Address', 'MAC Address', 'Device Type', 'Last Seen', 'Status'])
        for device in devices:
            writer.writerow([device['Agent ID'], device['Agent Name'], device['Device ID'], device['Device Name'], device['IP Address'], device['MAC Address'], device['Device Type'], device['Last Seen'], device['Status']])

def main():
    devices_to_delete = read_csv(INPUT_CSV)
    deleted_devices = []
    for device in devices_to_delete:
        try:
            if delete_device(device['Agent ID'], device['Device ID']):
                device['Status'] = 'Deleted'
            else:
                device['Status'] = 'Failed: Could not delete device'
        except requests.HTTPError as e:
            device['Status'] = f'Failed: {str(e)}'
        deleted_devices.append(device)
    save_to_csv(deleted_devices, OUTPUT_CSV)

if __name__ == '__main__':
    main()
