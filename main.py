import tkinter as tk
from tkinter import filedialog
import pandas as pd
import sqlite3
import xlrd
import xlwt


def select_file(title="Select a file"):
    return filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xls")])


def compare_neotech():
    # Prompt user to select last week's file
    last_week_file = select_file("Select last week's file")
    if not last_week_file:  # Check if a file was selected
        return

    # Prompt user to select current week's file
    current_week_file = select_file("Select the new file")
    if not current_week_file:  # Check if a file was selected
        return

    # Read the Excel files using xlrd
    last_week_data = pd.read_excel(last_week_file, sheet_name='Full File', engine='xlrd')
    current_week_data = pd.read_excel(current_week_file, sheet_name='Sheet1', engine='xlrd')

    # Convert column names to uppercase and strip white spaces
    last_week_data.columns = last_week_data.columns.str.upper().str.strip()
    current_week_data.columns = current_week_data.columns.str.upper().str.strip()

    print("Last week columns:", last_week_data.columns)
    print("Current week columns:", current_week_data.columns)

    # If 'PARTNUM' column not present in either dataframe, raise an error
    if 'PARTNUM' not in last_week_data.columns or 'PARTNUM' not in current_week_data.columns:
        raise ValueError("PartNum column not found in one of the files after adjustments.")

    # Process PARTNUM columns
    last_week_data['PARTNUM'] = last_week_data['PARTNUM'].astype(str).str.strip()
    current_week_data['PARTNUM'] = current_week_data['PARTNUM'].astype(str).str.strip()

    # Filter last week's data
    removed_from_prev = last_week_data[~last_week_data['PARTNUM'].isin(current_week_data['PARTNUM'])]

    # Create or connect to an SQLite database
    db_conn = sqlite3.connect('neotech_data.db')

    # Write data to SQLite database
    removed_from_prev.to_sql('removed_from_prev', db_conn, if_exists='replace', index=False)

    # Create a new workbook using xlwt
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('Removed from Prev File')

    # Write the column headers
    for col_idx, col_name in enumerate(removed_from_prev.columns):
        sheet.write(0, col_idx, col_name)

    # Write the data rows
    for row_idx, row in enumerate(removed_from_prev.values, start=1):
        for col_idx, value in enumerate(row):
            sheet.write(row_idx, col_idx, value)

    # Save the Excel file in .xls format
    result_file = filedialog.asksaveasfilename(defaultextension=".xls", filetypes=[("Excel files", "*.xls")])
    if result_file:
        book.save(result_file)
        print("File saved successfully:", result_file)
    else:
        print("File save canceled.")

    print("Process complete.")


# Create the GUI window
window = tk.Tk()

# Set the window geometry to a larger size
window.geometry("1200x730")

# Add a title label
title_label = tk.Label(window, text="Comparing Files For Neotech",
                       font=("Microsoft YaHei", 28, "bold", "underline"), foreground="red")
title_label.grid(row=0, column=0, pady=20)

# # Add instructions label
# instructions_label = tk.Label(window,
#                               text="Instructions:\n"
#                                    "To identify missing IPNs between the last and current RAW BOND files and add the \n"
#                                    "'Item_Type_Changed_To' and 'Sourced_Type_Changed_To' columns, follow these steps:\n"
#                                    "1. Select the Last RAW BOND Creation File.\n"
#                                    "2. Select the Current RAW BOND Creation File.",
#                               font=("Microsoft YaHei", 18))
# instructions_label.grid(row=1, column=0, pady=10)

# Create a frame for the first button
button_frame1 = tk.Frame(window)
button_frame1.grid(row=2, column=0, pady=10)

# Add a button to trigger file selection and comparison
compare_button = tk.Button(button_frame1, text='Compare NeoTech Files', command=compare_neotech,
                           font=("Microsoft YaHei", 20, "bold"), bg="red", fg="white")
compare_button.pack(fill='both')

# Configure the grid to expand
for i in range(6):
    window.grid_rowconfigure(i, weight=1)
window.grid_columnconfigure(0, weight=1)

# Run the GUI window
window.mainloop()
