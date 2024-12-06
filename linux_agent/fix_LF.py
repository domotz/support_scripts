#!/usr/bin/env python3

import os

def convert_to_unix_format(file_path):
    """
    Converts a file to Unix (LF) line endings.
    """
    try:
        with open(file_path, "rb") as file:
            content = file.read()

        # Replace CRLF (\r\n) with LF (\n)
        new_content = content.replace(b"\r\n", b"\n")

        # Write the converted content back to the file
        with open(file_path, "wb") as file:
            file.write(new_content)

        print(f"Converted: {file_path}")
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")

def process_directory(directory):
    """
    Processes all files in the directory (recursively) to ensure Unix (LF) line endings.
    """
    for root, _, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            convert_to_unix_format(file_path)

def main():
    """
    Main function to process files in the script's directory.
    """
    script_dir = os.path.dirname(os.path.realpath(__file__))
    print(f"Processing directory: {script_dir}")
    process_directory(script_dir)
    print("All files have been converted to Unix (LF) line endings.")

if __name__ == "__main__":
    main()
