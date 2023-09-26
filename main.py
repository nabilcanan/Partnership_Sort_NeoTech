import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import webbrowser
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment


def select_file(title="Select a file"):
    return filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xls")])


def format_headers_in_excel(filename):
    # Load the workbook
    workbook = load_workbook(filename)

    headers_to_color = ['PSOFT PART', 'PSID CT', 'QUOTED MFG', 'QUOTED PART', 'PART CLASS', 'AML CPN_MFGID', 'NAME', 'AML CPN_MFGNUM',
                        'AML CPN_MFGPARTNUM', 'LIFECYCLE STATUS']
    color_list = ["fffa9e", "fffa9e", "fffa9e", "fffa9e", "fffa9e", "f3f800", "f3f800", "f3f800", "f3f800", "f3f800"]
    header_color_mapping = dict(zip(headers_to_color, color_list))

    general_headers_to_color = ['COMPANY', 'PLANT', 'SITE NAME', 'PARTNUM', 'BUYERID', 'IUM', 'PUM', 'NCNR FLAG',
                                'COMM CODE', 'CUSTOMER PN', 'REFERENCE', 'VENDORNUM', 'VENDORID', 'VENDORNAME',
                                'CONSOLIDATED NAME', 'CONSOLIDATED NAME 2', 'GROUPDESC', 'MFGNUM', 'MFGID',
                                'MFGPARTNUM', 'BASEUNITPRICE', 'EFFECTIVEDATE', 'EXPIRATIONDATE', 'CLASSID',
                                'PURPOINT', 'PARTDESCRIPTION', 'PRODCODE', 'ACTIVE Y OR N', 'CONSOLIDATED CUST NAME',
                                'MRPSHARE_C', 'ASOF', 'MINORDERQTY', 'MFGLOTMULTIPLE', 'MINMFGLOTSIZE', 'LEADTIME',
                                'PRICELIST_LT', '90 DAY DEMAND', 'TOTAL DEMAND', 'SUMOFONHANDQTY']

    general_color = "00abee"  # Light blue

    align_wrap = Alignment(wrap_text=True)

    # Loop through all sheets in the workbook
    for sheet in workbook.worksheets:
        for cell in sheet[1]:  # Headers are in the first row
            cell.alignment = align_wrap
            # Color the specified headers in the "Full File Without Dupes" sheet
            if sheet.title == 'Full File Without Dupes':
                if cell.value in header_color_mapping:
                    cell.fill = PatternFill(start_color=header_color_mapping[cell.value],
                                            end_color=header_color_mapping[cell.value],
                                            fill_type="solid")
                elif cell.value in general_headers_to_color:
                    cell.fill = PatternFill(start_color=general_color, end_color=general_color, fill_type="solid")

    workbook.save(filename)


def save_to_excel(original_data, unique_data, removed_from_prev_data, output_name):
    with pd.ExcelWriter(output_name, engine='openpyxl') as writer:
        original_data.to_excel(writer, sheet_name='Full Original File', index=False)
        unique_data.to_excel(writer, sheet_name='Full File Without Dupes', index=False)
        removed_from_prev_data.to_excel(writer, sheet_name='Removed from Prev File', index=False)


def compare_neotech():
    # Prompt user to select current week's file
    current_week_file = select_file("Select the new file")
    if not current_week_file:
        print("Current week's file not selected!")
        return

    # Read the Excel file using pandas with xlrd engine
    current_week_data = pd.read_excel(current_week_file, sheet_name='Sheet1', engine='xlrd')
    current_week_data.columns = current_week_data.columns.str.upper().str.strip()

    if 'PARTNUM' not in current_week_data.columns:
        print("Error: 'PARTNUM' column not found in the current week's file.")
        return

    # Prompt user to select last week's file
    last_week_file = select_file("Select last week's file")
    if not last_week_file:
        print("Last week's file not selected!")
        return

    last_week_data = pd.read_excel(last_week_file, sheet_name='Full Original File', engine='xlrd')
    last_week_data.columns = last_week_data.columns.str.upper().str.strip()  # Convert last_week_data columns too

    # Read the 'Dupes Removed' sheet for merging
    prev_week_dupes_removed = pd.read_excel(last_week_file, sheet_name='Full File Without Dupes', engine='xlrd')
    prev_week_dupes_removed.columns = prev_week_dupes_removed.columns.str.upper().str.strip()
    print(prev_week_dupes_removed.columns)

    # Data from the 2nd file without duplicates
    unique_data = current_week_data.drop_duplicates(subset='PARTNUM', keep='first')

    # Merge data from prev_week_dupes_removed into unique_data for specified columns
    columns_to_merge = ['PARTNUM', 'PSOFT PART', 'PSID CT', 'QUOTED MFG', 'QUOTED PART', 'PART CLASS']
    unique_data = unique_data.merge(prev_week_dupes_removed[columns_to_merge], on='PARTNUM', how='left')

    # # Columns to drop, we don't want these columns because they are already created in our existing workbook and,
    # # We will create new columns that we VLOOKUP with our final file, we don't need these columns anymore
    # columns_to_drop_x = ['PSOFT PART', 'PSID CT_x', 'QUOTED MFG_x', 'QUOTED PART_x', 'PART CLASS_x', 'UNNAMED: 40']

    # unique_data.drop(columns=columns_to_drop_x, inplace=True)

    # PartNums present in the last week's file but missing in the "Full File Without Dupes"
    removed_from_prev_data = last_week_data[~last_week_data['PARTNUM'].isin(unique_data['PARTNUM'])]

    # Drop duplicates for the "Removed from Prev File"
    removed_from_prev_data = removed_from_prev_data.drop_duplicates(subset='PARTNUM', keep='first')

    # Choose the location and name for the output Excel file
    output_name = filedialog.asksaveasfilename(title="Save the final workbook", defaultextension=".xlsx",
                                               filetypes=[("Excel files", "*.xlsx")])

    # Save all the dataframes to the Excel file
    save_to_excel(current_week_data, unique_data, removed_from_prev_data, output_name)

    # Format the headers
    format_headers_in_excel(output_name)

    # Notify the user of a successful operation using a dialog box
    messagebox.showinfo("Success", "Operation completed successfully!")


# Create the GUI window
window = tk.Tk()

# Set the window geometry to a larger size
window.geometry("1200x500")


def open_readme_link():
    webbrowser.open('https://github.com/nabilcanan/Partnership_Sort_NeoTech/blob/main/README.md',
                    new=2)  # new=2 ensures the link opens in a new window.


# Add a title label
title_label = tk.Label(window, text="Comparing Files For Neotech",
                       font=("Microsoft YaHei", 28, "bold", "underline"), foreground="red")
title_label.grid(row=0, column=0, pady=20)

# Add instructions label
instructions_label = tk.Label(window,
                              text="Instructions:\n"
                                   "1. Select your most recent NeoTech Contract File.\n"
                                   "2. Select the previous Neotech Contract File.\n"
                                   "3. Finally choose where you'd like to save your final workbook\n",
                              font=("Microsoft YaHei", 18))
instructions_label.grid(row=1, column=0, pady=10)

# Create a frame for the first button
button_frame1 = tk.Frame(window)
button_frame1.grid(row=2, column=0, pady=10)

# Add a button to trigger file selection and comparison
compare_button = tk.Button(button_frame1, text='Compare NeoTech Files', command=compare_neotech,
                           font=("Microsoft YaHei", 20, "bold"), bg="red", fg="white")
compare_button.pack(fill='both')

# Create a frame for the README link button
button_frame2 = tk.Frame(window)
button_frame2.grid(row=3, column=0, pady=10)

readme_button = tk.Button(button_frame2, text='Open README', command=open_readme_link,
                          font=("Microsoft YaHei", 20), bg="blue", fg="white")
readme_button.pack(fill='both')

# Configure the grid to expand
for i in range(6):
    window.grid_rowconfigure(i, weight=1)
window.grid_columnconfigure(0, weight=1)

# Run the GUI window
window.mainloop()
