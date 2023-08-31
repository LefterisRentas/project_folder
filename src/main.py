import os
import tkinter as tk
from tkinter import filedialog  # Import filedialog
import openpyxl
import threading
import subprocess
import eel
import sys
import bottle

BUNDLE_TEMP_DIR = ''

try:
    if getattr(sys, 'frozen') and hasattr(sys, '_MEIPASS'):
        BUNDLE_TEMP_DIR = sys._MEIPASS
        bottle.TEMPLATE_PATH.insert(0, os.path.join(BUNDLE_TEMP_DIR, 'views'))
except:
    BUNDLE_TEMP_DIR = ''

exclude_strings = []

@eel.expose
def process_files(excel_file, amount):
    process_excel(excel_file, amount)

@eel.expose
def browse_excel():
    excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])

    # Update the entry field with the selected file
    eel.updateExcelEntry(excel_file)

@eel.expose
def browse_exclude():
    global exclude_strings
    exclude_file = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")])
    if exclude_file:
        with open(exclude_file, "r", encoding="utf-8") as f:
            exclude_strings = [line.strip() for line in f.readlines()]
        eel.updateExcludeEntry(exclude_file)
        eel.populateExcludeList(exclude_strings)


@eel.expose
def add_exclude_string(exclude_string):
    if exclude_string:
        if exclude_string not in exclude_strings:
            exclude_strings.append(exclude_string)
            eel.customInsert(exclude_string)
            eel.clearExcludeEntry()
        else:
            eel.show_message("Duplicate Entry", "The exclude string already exists.")

@eel.expose
def remove_selected_exclude(selected_items):
    for selected_item in selected_items:
        try:
            exclude_strings.remove(selected_item)
        except ValueError:
            # Handle case where the item is not in the list
            pass
        eel.customRemove(selected_item)


@eel.expose
def save_changes():
    global exclude_strings

    # Get the file path from the entry field using the eel API
    exclude_file_path = eel.get_exclude_file()()
    directory = os.path.dirname(exclude_file_path)
    # If the entry is empty or the file doesn't exist, create a new file
    if not os.path.exists(directory):
        os.makedirs(os.getcwd().join("execute.txt"))

    try:
        with open(exclude_file_path, "w", encoding="utf-8") as f:
            for exclude_string in exclude_strings:
                f.write(exclude_string + "\n")

        eel.updateExcludeEntry(exclude_file_path)
        exclude_strings.clear()  # Clear the list in Python
        eel.clearExcludeEntry()  # Clear the entry in the UI

        # Reopen the file to populate the entry field again
        with open(exclude_file_path, "r", encoding="utf-8") as f:
            exclude_strings = [line.strip() for line in f.readlines()]
        eel.populateExcludeList(exclude_strings)  # Update the list in the UI
    except Exception as e:
        eel.show_message("Error", str(e))  # Display error message







def process_excel(excel_file, amount):
    global exclude_strings
    try:
        # Load the Excel workbook
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active

        # Identify column indices based on headers
        header_row = 6
        column_mapping = {}
        for col_idx, cell in enumerate(sheet[header_row], start=1):
            header = cell.value
            if header == "Συνολική":
                column_mapping["total"] = col_idx
            elif header == "Διεύθυνση":
                column_mapping["address"] = col_idx
            elif header == "Επωνυμία":
                column_mapping["name"] = col_idx

        # Calculate totals and exclude orders
        address_totals = {}  # Nested dictionary: address -> name -> total
        for row in sheet.iter_rows(min_row=header_row + 1):
            total_value = row[column_mapping["total"] - 1].value
            address = row[column_mapping["address"] - 1].value
            name = row[column_mapping["name"] - 1].value

            if total_value is not None and address is not None and name is not None:
                if name not in exclude_strings:
                    address_totals.setdefault(address, {})
                    address_totals[address].setdefault(name, 0)
                    address_totals[address][name] += total_value

        output_data = []
        for address, name_totals in address_totals.items():
            for name, total in name_totals.items():
                if total <= amount:
                    output_data.append((name, address, total))

        # Create and save the output Excel file
        output_workbook = openpyxl.Workbook()
        output_sheet = output_workbook.active
        output_sheet.append(["Επωνυμία", "Διεύθυνση", "Συνολική"])
        for entry in output_data:
            output_sheet.append(entry)
        output_workbook.save("output.xlsx")

        # Open the output.xlsx file with its default application
        subprocess.call(["start", "output.xlsx"], shell=True)
    except Exception as e:
        eel.show_message("Error", str(e))

eel.init(os.getcwd())
eel.start("main.html", size=(700, 1040),port=3003, host='localhost')
