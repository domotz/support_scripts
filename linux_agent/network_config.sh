#!/bin/bash
# Netplan Configurator script for Linux
# What it does:
# - Configures network interfaces on a Linux host using netplan

ver="1.0"

# Check if the script is run as root
if [[ $EUID -ne 0 ]]; then
    echo "This script must be run as root"
    exit 1
fi

echo "+----------------------------------------------------------+"
echo "|         ___                             _                |"
echo "|        (  _'\                          ( )_              |"
echo "|        | | ) |   _     ___ ___     _   | ,_) ____        |"
echo "|        | | | ) /'_'\ /' _ ' _ '\ /'_'\ | |  (_  ,)       |"
echo "|        | |_) |( (_) )| ( ) ( ) |( (_) )| |_  /'/_        |"
echo "|        (____/''\___/'(_) (_) (_)'\___/''\__)(____)       |"
echo "| -------------------------------------------------------- |"
echo "| The IT Infrastructure Monitoring and Management Solution |"
echo "+---------------------------------------------------------+"
echo ""
echo "Netplan Configurator for Ubuntu Linux ver ${ver}"
echo "========================"
echo "Before using this, make sure your system uses the"
echo "netplan file to configure its network properties"
echo "========================"

# Change this to your netplan file path
NETPLAN_FILE="/etc/netplan/00-installer-config.yaml"

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
            printf "Invalid IP address or subnet mask wider than /16. Please enter a valid IP address.\n" >&2
        fi
    done
    echo "========================"
    
    while true; do
        read -p "Enter the gateway IP address: " gateway
        if validate_ip "$gateway"; then
            break
        else
            printf "Invalid IP address. Please enter a valid gateway IP address.\n" >&2
        fi
    done
    echo "========================"

    while true; do
        read -p "Enter the first DNS server IP address: " dns1
        if validate_ip "$dns1"; then
            break
        else
            printf "Invalid IP address. Please enter a valid DNS server IP address.\n" >&2
        fi
    done
    echo "========================"

    while true; do
        read -p "Enter the second DNS server IP address: " dns2
        if validate_ip "$dns2"; then
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
    while [[ -f "$backup_dir/${backup_name}.bak$i" ]]; do
        let i++
    done
    cp "$file" "$backup_dir/${backup_name}.bak$i"
    echo "Backup made as ${backup_name}.bak$i"
}
echo "========================"

# Preview the netplan file with new configurations
echo "Preview of the netplan file with new configurations:"

# Read the existing netplan file, excluding any existing interface sections
if ! awk '
    /^    '"${IFACE}"':/ {delete_block=1}
    /^[^ ]/ {delete_block=0}
    !delete_block' "$NETPLAN_FILE" > "${NETPLAN_FILE}.tmp"; then
    printf "Error reading netplan file.\n" >&2
    exit 1
fi

# Append the new interface configuration
cat <<EOL >> "${NETPLAN_FILE}.tmp"
    ${IFACE}:
      dhcp4: ${dhcp}
EOL

if [[ $dhcp == "no" ]]; then
    cat <<EOL >> "${NETPLAN_FILE}.tmp"
      addresses:
        - ${ip}
      routes:
        - to: 0.0.0.0/0
          via: ${gateway}
      nameservers:
        addresses:
          - ${dns1}
          - ${dns2}
EOL
fi

# Append any VLAN configurations
awk '
    /^  vlans:/ {print; getline; while ($0 ~ /^    /) {print; getline}}
' "$NETPLAN_FILE" >> "${NETPLAN_FILE}.tmp"

cat "${NETPLAN_FILE}.tmp"
echo "========================"

# Confirmation
read -p "Do you want to apply these settings? (yes/no): " confirm
if [[ $confirm == "yes" ]]; then
    make_backup "$NETPLAN_FILE"
    mv "${NETPLAN_FILE}.tmp" "$NETPLAN_FILE"
    echo "Configurations have been written to $NETPLAN_FILE. Please review and apply with 'netplan apply'."
else
    echo "Changes not applied."
    rm "${NETPLAN_FILE}.tmp"
    exit 0
fi
echo "========================"