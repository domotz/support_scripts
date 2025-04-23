import requests
import csv
import datetime

API_KEY = 'YOUR_API_KEY'  # Replace with your actual API key
API_URL = 'https://api-us-east-1-cell-1.domotz.com/public-api/v1'
OFFLINE_THRESHOLD_DAYS = 14

def get_agents():
    response = requests.get(f'{API_URL}/agent', headers={'x-api-key': API_KEY})
    response.raise_for_status()
    return response.json()

def get_devices(agent_id):
    response = requests.get(f'{API_URL}/agent/{agent_id}/device', headers={'x-api-key': API_KEY})
    response.raise_for_status()
    return response.json()

def filter_offline_devices(devices):
    offline_devices = []
    now = datetime.datetime.now(datetime.timezone.utc)
    threshold_date = now - datetime.timedelta(days=OFFLINE_THRESHOLD_DAYS)
    for device in devices:
        if 'last_status_change' in device and device['status'] == 'DOWN':
            try:
                last_seen = datetime.datetime.strptime(device['last_status_change'], '%Y-%m-%dT%H:%M:%S%z')
            except ValueError:
                last_seen = datetime.datetime.strptime(device['last_status_change'], '%Y-%m-%dT%H:%M:%S.%f%z')
            if last_seen < threshold_date:
                offline_devices.append(device)
    return offline_devices

def save_to_csv(devices, filename='offline_devices.csv'):
    with open(filename, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Agent ID', 'Agent Name', 'Device ID', 'Device Name', 'IP Address', 'MAC Address', 'Device Type', 'Last Seen'])
        for device in devices:
            device_type = device.get('type', {}).get('label', 'Unknown')
            hw_address = device.get('hw_address', 'Unknown')
            ip_addresses = ', '.join(device.get('ip_addresses', []))
            writer.writerow([device['agent_id'], device['agent_name'], device['id'], device['display_name'], ip_addresses, hw_address, device_type, device['last_status_change']])

def main():
    all_offline_devices = []
    agents = get_agents()
    for agent in agents:
        agent_id = agent['id']
        agent_name = agent['display_name']
        devices = get_devices(agent_id)
        offline_devices = filter_offline_devices(devices)
        for device in offline_devices:
            device['agent_id'] = agent_id
            device['agent_name'] = agent_name
        all_offline_devices.extend(offline_devices)
    save_to_csv(all_offline_devices)

if __name__ == '__main__':
    main()
