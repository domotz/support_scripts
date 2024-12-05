#!/bin/bash
# Domotz script for downloading the configuration scripts
# What it does:
# - download script to configure a Linux System for Domotz Monitoring  

echo "+------------------------------------------------+"
echo "|  ___                             _             |"
echo "| (  _'\                          ( )_           |"
echo "| | | ) |   _     ___ ___     _   | ,_) ____     |"
echo "| | | | ) /'_'\ /' _ ' _ '\ /'_'\ | |  (_  ,)    |"
echo "| | |_) |( (_) )| ( ) ( ) |( (_) )| |_  /'/_     |"
echo "| (____/''\___/'(_) (_) (_)'\___/''\__)(____)    |"
echo "| ---------------------------------------------- |"
echo "| The IT Monitoring and Management Solition      |"
echo "+------------------------------------------------+"
echo "=================================================="
echo "Domotz Script to download updated scripts to "
echo "configure a Linux system for Domotz Monitoring"
echo "=================================================="

set -euo pipefail

# Global variables
REPO_URL="https://github.com/domotz/support_scripts/raw/refs/heads/develop/linux_agent/domotz_linux_scripts.tar.gz"
TEMP_DIR="/tmp/repo_download"

# Function to validate dependencies
validate_dependencies() {
    local dependencies=("curl" "tar")
    for cmd in "${dependencies[@]}"; do
        if ! command -v "$cmd" &>/dev/null; then
            printf "Error: Required command '%s' is not installed.\n" "$cmd" >&2
            return 1
        fi
    done
}

# Function to download the file
download_file() {
    local url="$1"
    local output_dir="$2"

    local filename; filename=$(basename "$url")
    local output_path="$output_dir/$filename"

    printf "Downloading file from %s...\n" "$url" >&2
    if ! curl -fSL "$url" -o "$output_path"; then
        printf "Error: Failed to download file from %s\n" "$url" >&2
        return 1
    fi

    if [[ ! -s "$output_path" ]]; then
        printf "Error: Downloaded file is empty or corrupted: %s\n" "$output_path" >&2
        return 1
    fi

    printf "%s\n" "$output_path"  # Return only the file path
}

# Function to extract the archive
extract_archive() {
    local archive_path="$1"
    local destination="$2"

    printf "Extracting archive %s to %s...\n" "$archive_path" "$destination" >&2
    if ! tar -xzf "$archive_path" -C "$destination"; then
        printf "Error: Failed to extract archive %s\n" "$archive_path" >&2
        return 1
    fi
    printf "Extraction complete.\n" >&2
}

# Main function
main() {
    validate_dependencies

    # Prepare temporary directory
    mkdir -p "$TEMP_DIR"

    local tar_path
    tar_path=$(download_file "$REPO_URL" "$TEMP_DIR")

    # Extract to the current working directory
    extract_archive "$tar_path" "$(pwd)"

    # Cleanup
    rm -rf "$TEMP_DIR"

    printf "Operation completed successfully.\n"
}

main "$@"
