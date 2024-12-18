#!/bin/bash
# Description: This script downloads and extracts the latest Domotz configuration scripts
#              for setting up a Protectli Domotz system for Domotz monitoring.
#
# Features:
# - Downloads the latest Domotz scripts as a compressed archive from a GitHub repository.
# - Extracts the archive to the directory where the script is executed.
# - Validates dependencies (e.g., curl, tar) before execution.
# - Provides error handling for failed downloads or extraction issues.
# - Displays a Domotz-themed banner to highlight its purpose.
#
# Usage:
# - Run the script with sufficient privileges (e.g., sudo) to ensure access to necessary resources.
# - The extracted files will appear in the same directory where the script is executed.
#
# Requirements:
# - Linux-based system with the following installed:
#   - curl: For downloading files from the internet.
#   - tar: For extracting compressed archives.
#
# Global Variables:
# - REPO_URL: URL pointing to the latest Domotz configuration scripts archive.
# - TEMP_DIR: Temporary directory used for downloading files before extraction.
#
# Note: Ensure that all dependencies are installed before running the script.
#       If issues persist, check internet connectivity and the validity of the REPO_URL.
# 
# ChangeLog
# 1.0 -- Created on 6 Dec 2024
ver=1.0
set -euo pipefail

echo "+------------------------------------------------+"
echo "|  ___                             _             |"
echo "| (  _'\                          ( )_           |"
echo "| | | ) |   _     ___ ___     _   | ,_) ____     |"
echo "| | | | ) /'_'\ /' _ ' _ '\ /'_'\ | |  (_  ,)    |"
echo "| | |_) |( (_) )| ( ) ( ) |( (_) )| |_  /'/_     |"
echo "| (____/''\___/'(_) (_) (_)'\___/''\__)(____)    |"
echo "| ---------------------------------------------- |"
echo "| The IT Monitoring and Management Solution      |"
echo "+------------------------------------------------+"
echo "=================================================="
echo "Domotz ${ver} script to get latest tools to     "
echo "configure a Protectli device for Domotz Monitoring"
echo "=================================================="
read -p "Please press Enter to continue: "


# Global variables
REPO_URL="https://github.com/domotz/support_scripts/raw/refs/heads/main/linux_agent/domotz_protectli_scripts_latest.tar.gz"
TEMP_DIR="/tmp/domotz_download"
EXTRACTED_DIR_NAME="domotz_protectli_scripts_latest"

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
    if ! wget -q -O "$output_path" "$url"; then
        printf "Error: Failed to download file from %s\n" "$url" >&2
        return 1
    fi

    if [[ ! -s "$output_path" ]]; then
        printf "Error: Downloaded file is empty or corrupted: %s\n" "$output_path" >&2
        return 1
    fi

    printf "%s\n" "$output_path"  # Return only the file path
}


# Function to remove existing directory
remove_old_directory() {
    local directory="$1"

    if [[ -d "$directory" ]]; then
        printf "Removing existing directory: %s\n" "$directory"
        if ! rm -rf "$directory"; then
            printf "Error: Failed to remove directory %s\n" "$directory" >&2
            return 1
        fi
    fi
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

    # Remove old extracted directory
    remove_old_directory "$(pwd)/$EXTRACTED_DIR_NAME"

    # Download the archive
    local tar_path
    tar_path=$(download_file "$REPO_URL" "$TEMP_DIR")

    # Extract to the current working directory
    extract_archive "$tar_path" "$(pwd)"

    # Cleanup
    rm -rf "$TEMP_DIR"

    printf "Operation completed successfully.\n"
}

main "$@"
