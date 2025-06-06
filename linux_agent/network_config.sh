#!/bin/bash
# This script provides an example configuration for setting up and configure the network on a Linux system.
# Prerequisites:
# - the system should use the netplan file for network configuration.
ver="1.1"

# !!!Please changhe the path of the netplan file according to your system setup.!!!

# Change this to your netplan file path
NETPLAN_FILE="/etc/netplan/00-installer-config.yaml"


# Check if the script is run as root
if [[ $EUID -ne 0 ]]; then
    echo "This script must be run as root"
    exit 1
fi

# Function to check if the IP address is valid
validate_ip() {
    local ip_with_mask=$1
    local ip=${ip_with_mask%%/*}
    local mask=${ip_with_mask##*/}

    # Validate IP format
    if [[ $ip =~ ^([0-9]{1,3}\.){3}[0-9]{1,3}$ ]]; then
        IFS='.' read -r -a octets <<< "$ip"
        for octet in "${octets[@]}"; do
            if ((octet < 0 || octet > 255)); then
                return 1
            fi
        done
    else
        return 1
    fi

    # Validate subnet mask if present
    if [[ $ip_with_mask == */* ]]; then
        # Netplan typically expects CIDR prefix length (e.g., 24 for /24)
        # The original script's check ((mask < 16 || mask > 32)) is specific.
        if ! [[ $mask =~ ^[0-9]{1,2}$ ]] || ((mask < 16 || mask > 32)); then
            return 1
        fi
    fi

    return 0
}

# Function to add /24 subnet mask if not provided
add_subnet_mask() {
    local ip=$1
    if [[ ! $ip =~ /[0-9]{1,2}$ ]]; then
        ip="${ip}/24"
    fi
    echo "$ip"
}

# Function to check if the input is a number
is_number() {
    local num=$1
    [[ $num =~ ^[0-9]+$ ]]
}

# Show the interfaces currently configured
echo "Current interfaces configuration:"
echo ""
ip -o link show | awk -F': ' '{print $2}' | grep -v lo
echo ""
echo "========================"

# Function to check if an interface exists
validate_interface() {
    local iface=$1
    ip -o link show | awk -F': ' '{print $2}' | grep -v lo | grep -wq "$iface"
}

# Ask which interface to configure
while true; do
    read -p "Enter the interface name to configure: " IFACE
    if validate_interface "$IFACE"; then
        break
    else
        printf "Interface not found. Please enter a valid interface name.\n" >&2
    fi
done
echo "========================"

# Ask for DHCP or Static IP
while true; do
    read -p "Do you want to use DHCP for the interface? (yes/no): " dhcp
    if [[ $dhcp == "yes" || $dhcp == "no" ]]; then
        break
    else
        printf "Invalid input. Please enter 'yes' or 'no'.\n" >&2
    fi
done
echo "========================"

# If Static IP, get the details
if [[ $dhcp == "no" ]]; then
    while true; do
        read -p "Enter the static IP address (e.g., 192.168.1.10 or 192.168.1.10/24): " ip
        ip=$(add_subnet_mask "$ip")
        if validate_ip "$ip"; then
            break
        else
            printf "Invalid IP address or subnet mask not in range /16-/32. Please enter a valid IP address.\n" >&2
        fi
    done
    echo "========================"

    while true; do
        read -p "Enter the gateway IP address: " gateway
        if validate_ip "${gateway}"; then # Validate basic IP format for gateway
            break
        else
            printf "Invalid IP address. Please enter a valid gateway IP address.\n" >&2
        fi
    done
    echo "========================"

    while true; do
        read -p "Enter the first DNS server IP address: " dns1
        if validate_ip "${dns1}"; then # Validate basic IP format for DNS
            break
        else
            printf "Invalid IP address. Please enter a valid DNS server IP address.\n" >&2
        fi
    done
    echo "========================"

    while true; do
        read -p "Enter the second DNS server IP address: " dns2
        if validate_ip "${dns2}"; then # Validate basic IP format for DNS
            break
        else
            printf "Invalid IP address. Please enter a valid DNS server IP address.\n" >&2
        fi
    done
    echo "========================"
fi

# Define a function to make backup with incrementing names
make_backup() {
    local file=$1
    local backup_dir; backup_dir=$(dirname "$file")
    local backup_name; backup_name=$(basename "$file")
    local i=1

    # Check if the original file exists before trying to backup
    if [[ ! -f "$file" ]]; then
        echo "Original file $file does not exist. Skipping backup."
        return
    fi

    while [[ -f "$backup_dir/${backup_name}.bak$i" ]]; do
        let i++
    done
    cp "$file" "$backup_dir/${backup_name}.bak$i"
    echo "Backup made as ${backup_name}.bak$i"
}
echo "========================" # Separator after make_backup function definition

# ---- MODIFICATIONS ACCORDING TO USER REQUEST ----
# Initialize the temporary netplan file by writing the header.
# This effectively clears any previous content for NETPLAN_FILE.tmp
# and starts it with the required network configuration structure.
cat <<EOL > "${NETPLAN_FILE}.tmp"
network:
  version: 2
  ethernets:
EOL

# Append the new interface configuration to the initialized temporary file
# The indentation (4 spaces) ensures it's correctly placed under 'ethernets:'
cat <<EOL >> "${NETPLAN_FILE}.tmp"
    ${IFACE}:
      dhcp4: ${dhcp}
EOL

if [[ $dhcp == "no" ]]; then
    cat <<EOL >> "${NETPLAN_FILE}.tmp"
      addresses:
        - ${ip}
      routes:
        - to: default # Using "default" is more common in Netplan than 0.0.0.0/0
          via: ${gateway}
      nameservers:
        addresses:
          - ${dns1}
          - ${dns2}
EOL
fi
# ---- END OF MODIFICATIONS ----


# Preview the netplan file with new configurations:
echo "Preview of the netplan file with new configurations:"

# The old awk commands to read from NETPLAN_FILE and to append other configurations (like VLANs)
# have been removed as per the requirement to clear the file and start fresh with the specified header.

cat "${NETPLAN_FILE}.tmp"
echo "========================"

# Confirmation
read -p "Do you want to apply these settings? (yes/no): " confirm
if [[ $confirm == "yes" ]]; then
    make_backup "$NETPLAN_FILE"
    # Move the newly constructed temporary file to the actual netplan file path.
    # This overwrites the original file, effectively "clearing" it and applying the new content.
    if mv "${NETPLAN_FILE}.tmp" "$NETPLAN_FILE"; then
        echo "Configurations have been written to $NETPLAN_FILE. Please review and apply with 'sudo netplan apply'."
    else
        echo "Error: Failed to write configurations to $NETPLAN_FILE." >&2
        if [[ -f "${NETPLAN_FILE}.tmp" ]]; then
             echo "The new configuration is still available in ${NETPLAN_FILE}.tmp" >&2
        fi
        exit 1
    fi
else
    echo "Changes not applied."
    rm -f "${NETPLAN_FILE}.tmp" # Use -f to suppress error if file doesn't exist for rm
    exit 0
fi
echo "========================"