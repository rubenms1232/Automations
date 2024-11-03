import os
import webbrowser
import time
from pathlib import Path
import sys
import argparse
import tkinter as tk
from tkinter import filedialog

def select_folder_dialog():
    """Open a folder selection dialog"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    folder_path = filedialog.askdirectory(title="Select Folder with Documents to Print")
    return folder_path if folder_path else None

def get_folder_path():
    """Get folder path from arguments or user input"""
    parser = argparse.ArgumentParser(description='Batch print documents from a folder')
    parser.add_argument('--path', help='Path to folder containing documents')
    args = parser.parse_args()

    folder_path = args.path
    
    if not folder_path:
        print("\nNo folder path provided in arguments.")
        print("1. Enter folder path manually")
        print("2. Use file dialog to select folder")
        print("3. Use current directory")
        
        choice = input("\nEnter your choice (1-3): ").strip()
        
        if choice == '1':
            folder_path = input("\nEnter the folder path: ").strip()
        elif choice == '2':
            folder_path = select_folder_dialog()
            if not folder_path:
                print("No folder selected. Using current directory.")
                folder_path = os.getcwd()
        else:
            folder_path = os.getcwd()

    # Validate and convert to absolute path
    try:
        folder_path = os.path.abspath(folder_path)
        if not os.path.isdir(folder_path):
            print(f"Error: '{folder_path}' is not a valid directory")
            sys.exit(1)
    except Exception as e:
        print(f"Error with folder path: {e}")
        sys.exit(1)
        
    return folder_path

def get_printable_files(folder_path):
    """Get list of printable files from the folder"""
    printable_extensions = {'.pdf', '.doc', '.docx', '.txt', '.html', '.htm'}
    files = []
    
    try:
        for file in Path(folder_path).iterdir():
            if file.is_file() and file.suffix.lower() in printable_extensions:
                files.append(file)
    except Exception as e:
        print(f"Error accessing folder: {e}")
        sys.exit(1)
        
    return files

def open_files_in_batches(files, batch_size=5):
    """Open files in browser in batches"""
    total_files = len(files)
    
    if total_files == 0:
        print("No printable files found in the folder!")
        return
        
    for i in range(0, total_files, batch_size):
        batch = files[i:i + batch_size]
        
        print(f"\nOpening batch {i//batch_size + 1} of {(total_files + batch_size - 1)//batch_size}:")
        for file in batch:
            print(f"- {file.name}")
            try:
                file_url = file.absolute().as_uri()
                webbrowser.open(file_url, new=2)
                time.sleep(1)  # Small delay between opening files
            except Exception as e:
                print(f"Error opening {file}: {e}")
        
        # Wait for user to process current batch
        if i + batch_size < total_files:
            input(f"\nPress Enter when ready for next batch...")

def main():
    print("=== Batch Document Printer ===")
    
    # Get folder path using the new function
    folder_path = get_folder_path()
    
    print(f"\nScanning folder: {folder_path}")
    files = get_printable_files(folder_path)
    
    print(f"Found {len(files)} printable files")
    if files:
        print("\nFiles will be opened in batches of 5.")
        print("Each batch will open in your default web browser.")
        print("You can manually print each file and close the tabs.")
        input("Press Enter to start...")
        
        open_files_in_batches(files)
        
        print("\nAll files have been processed!")

if __name__ == "__main__":
    main()