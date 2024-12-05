#!/bin/bash
# Description: This script enables or disables the SSH service on an Ubuntu 22.04 system.
#              It ensures proper management of the SSH service via systemd and updates 
#              the UFW firewall rules to allow or deny SSH connections accordingly.
# 
# Usage:       ./manage_ssh.sh <enable|disable>
#              - "enable"  : Starts and enables the SSH service, and allows SSH in the UFW firewall.
#              - "disable" : Stops and disables the SSH service, and removes SSH allowance from the firewall.
# 
# Requirements:
#              - Ubuntu 22.04 or a compatible system
#              - sudo privileges
#              - systemd for service management
#              - UFW (Uncomplicated Firewall) installed and configured
set -euo pipefail

# Functions to validate dependencies
validate_dependencies() {
    local dependencies=("systemctl" "ufw")
    for cmd in "${dependencies[@]}"; do
        if ! command -v "$cmd" &>/dev/null; then
            printf "Error: Required command '%s' is not installed.\n" "$cmd" >&2
            return 1
        fi
    done
}

# Function to enable SSH service
enable_ssh() {
    printf "Enabling SSH service...\n"
    if ! sudo systemctl start ssh && sudo systemctl enable ssh; then
        printf "Error: Failed to enable SSH service.\n" >&2
        return 1
    fi

    printf "Allowing SSH through UFW firewall...\n"
    if ! sudo ufw allow ssh; then
        printf "Error: Failed to allow SSH through the firewall.\n" >&2
        return 1
    fi

    printf "SSH service has been enabled, and firewall updated.\n"
}

# Function to disable SSH service
disable_ssh() {
    printf "Disabling SSH service...\n"
    if ! sudo systemctl stop ssh && sudo systemctl disable ssh; then
        printf "Error: Failed to disable SSH service.\n" >&2
        return 1
    fi

    printf "Removing SSH allowance from UFW firewall...\n"
    if ! sudo ufw delete allow ssh; then
        printf "Error: Failed to remove SSH allowance from the firewall.\n" >&2
        return 1
    fi

    printf "SSH service has been disabled, and firewall updated.\n"
}

# Main function
main() {
    validate_dependencies

    if [[ $# -ne 1 ]]; then
        printf "Usage: %s <enable|disable>\n" "$(basename "$0")" >&2
        return 1
    fi

    local action="$1"
    case "$action" in
        enable)
            enable_ssh
            ;;
        disable)
            disable_ssh
            ;;
        *)
            printf "Invalid argument: %s. Use 'enable' or 'disable'.\n" "$action" >&2
            return 1
            ;;
    esac
}

main "$@"
