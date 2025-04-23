import requests
import csv

API_KEY = 'YOUR_API_KEY'  # Replace with your actual API key
API_URL = 'https://api-us-east-1-cell-1.domotz.com/public-api/v1'

def get_agents():
    response = requests.get(f'{API_URL}/agent', headers={'x-api-key': API_KEY})
    response.raise_for_status()
    return response.json()

def get_devices(agent_id):
    response = requests.get(f'{API_URL}/agent/{agent_id}/device', headers={'x-api-key': API_KEY})
    response.raise_for_status()
    return response.json()

def find_discrepancies(devices):
    discrepancies = []
    for device in devices:
        name = device.get('display_name', '')
        dhcp_name = device.get('names', {}).get('dhcp', '')
        if name != dhcp_name:
            discrepancies.append({
                'Agent ID': device['agent_id'],
                'Agent Name': device['agent_name'],
                'Device ID': device['id'],
                'Device Name': name,
                'DHCP Name': dhcp_name
            })
    return discrepancies

def save_to_csv(devices, filename='device_name_discrepancies.csv'):
    with open(filename, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Agent ID', 'Agent Name', 'Device ID', 'Device Name', 'DHCP Name'])
        for device in devices:
            writer.writerow([device['Agent ID'], device['Agent Name'], device['Device ID'], device['Device Name'], device['DHCP Name']])

def main():
    all_discrepancies = []
    agents = get_agents()
    for agent in agents:
        agent_id = agent['id']
        agent_name = agent['display_name']
        devices = get_devices(agent_id)
        for device in devices:
            device['agent_id'] = agent_id
            device['agent_name'] = agent_name
        discrepancies = find_discrepancies(devices)
        all_discrepancies.extend(discrepancies)
    save_to_csv(all_discrepancies)

if __name__ == '__main__':
    main()
