import tkinter as tk
from tkinter import filedialog, messagebox
import numpy as np
import pandas as pd
import sqlite3
import xlwt


def select_file(title="Select a file"):
    return filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xls")])


def convert_dtype(value):
    if isinstance(value, np.generic):
        return value.item()
    return value


def combine_workbooks_with_xlwt(target_workbook, new_sheet_data, new_sheet_name, output_name):
    with pd.ExcelFile(target_workbook, engine='xlrd') as xls:
        sheet_names = xls.sheet_names

        book = xlwt.Workbook(encoding='utf-8')

        # Write each sheet from the original workbook to the new workbook
        for sheet in sheet_names:
            data = pd.read_excel(xls, sheet_name=sheet, engine='xlrd')

            # Add sheet to workbook
            ws = book.add_sheet(sheet)

            # Write headers
            for col_idx, col in enumerate(data.columns):
                ws.write(0, col_idx, col)

            # Write data
            for row_idx, index in enumerate(data.index):
                for col_idx, col in enumerate(data.columns):
                    value = data.at[index, col]

                    # Convert numpy data types to native Python types
                    value = convert_dtype(value)

                    if not pd.isna(value):
                        ws.write(row_idx + 1, col_idx, value)

        # Add the new sheet data to the new workbook
        ws = book.add_sheet(new_sheet_name)

        # Write headers for new sheet
        for col_idx, col in enumerate(new_sheet_data.columns):
            ws.write(0, col_idx, col)

        # Write data for new sheet
        for row_idx, index in enumerate(new_sheet_data.index):
            for col_idx, col in enumerate(new_sheet_data.columns):
                value = new_sheet_data.at[index, col]

                # Convert numpy data types to native Python types
                value = convert_dtype(value)

                if not pd.isna(value):
                    ws.write(row_idx + 1, col_idx, value)

        # Save the new workbook
        book.save(output_name)


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

    # If 'PARTNUM' column not present in either dataframe, raise an error
    if 'PARTNUM' not in last_week_data.columns or 'PARTNUM' not in current_week_data.columns:
        raise ValueError("PartNum column not found in one of the files after adjustments.")

    # Process PARTNUM columns
    last_week_data['PARTNUM'] = last_week_data['PARTNUM'].astype(str).str.strip()
    current_week_data['PARTNUM'] = current_week_data['PARTNUM'].astype(str).str.strip()

    # Filter last week's data
    removed_from_prev = last_week_data[~last_week_data['PARTNUM'].isin(current_week_data['PARTNUM'])]

    # Create or connect to an SQLite database (this step is kept from your original code, modify if needed)
    db_conn = sqlite3.connect('neotech_data.db')
    removed_from_prev.to_sql('removed_from_prev', db_conn, if_exists='replace', index=False)

    # Ask the user to select the workbook into which the new sheet will be added
    target_workbook = select_file("Choose the workbook where you want to add the 'Removed from Prev File' sheet")
    if not target_workbook:
        print("No workbook selected to add the sheet.")
        return

    # Let user specify the output name for the combined workbook
    output_name = filedialog.asksaveasfilename(title="Save the combined workbook", defaultextension=".xls",
                                               filetypes=[("Excel files", "*.xls")])
    if not output_name:
        print("File save canceled.")
        return

    # Combine workbooks using the new method
    combine_workbooks_with_xlwt(target_workbook, removed_from_prev, "Removed From Prev File", output_name)

    print("Sheet 'Removed From Prev File' added successfully to", output_name)
    print("Process complete.")
    # Show success message
    messagebox.showinfo("Congrats You're A Genius!", "Success Final Workbook Saved")


# Create the GUI window
window = tk.Tk()

# Set the window geometry to a larger size
window.geometry("1200x500")

# Add a title label
title_label = tk.Label(window, text="Comparing Files For Neotech",
                       font=("Microsoft YaHei", 28, "bold", "underline"), foreground="red")
title_label.grid(row=0, column=0, pady=20)

# Add instructions label
instructions_label = tk.Label(window,
                              text="Instructions:\n"
                                   "1. Select your previous NeoTech Contract File.\n"
                                   "2. Choose the most recent NeoTech contract File.\n"
                                   "3. Choose the most recent NeoTech contract File, this is where you're adding the sheet. (Step 2)\n"
                                   "4. Finally choose where you'd like to save your final workbook\n",
                              font=("Microsoft YaHei", 18))
instructions_label.grid(row=1, column=0, pady=10)

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
