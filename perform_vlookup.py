from tkinter import filedialog, messagebox
import pandas as pd


def select_file(title="Select a file"):
    return filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xls")])


def perform_vlookup():
    # Ask the user to select the file that contains the data they want to merge (VLOOKUP data source)
    lookup_file = select_file("Select the file with data for VLOOKUP")

    # Check if a lookup file is selected
    if not lookup_file:
        messagebox.showerror("Error", "VLOOKUP data source file not selected!")
        return

    # Ask the user to select the target file where the 'Dupes Removed' sheet is located and where the VLOOKUP needs
    # to be performed
    target_file = select_file("Select the file where you need to perform VLOOKUP (contains 'Dupes Removed' sheet)")

    # Check if a target file is selected
    if not target_file:
        messagebox.showerror("Error", "Target file not selected!")
        return

    # Load data from both files into pandas DataFrames
    lookup_data = pd.read_excel(lookup_file)
    target_data = pd.read_excel(target_file, sheet_name='Dupes Removed')
    original_data = pd.read_excel(target_file, sheet_name='Original Data')
    removed_data = pd.read_excel(target_file, sheet_name='Lost Items')

    # Rename 'PartNum' column in lookup_data to 'PARTNUM' for consistency
    lookup_data = lookup_data.rename(columns={'PartNum': 'PARTNUM'})

    # List of columns to bring from the lookup file
    columns_to_merge = ['PSoft Part', 'PSID CT', 'Quoted Mfg', 'Quoted Part', 'Part Class']

    # Filter only the columns we need from the lookup data
    lookup_data_filtered = lookup_data[['PARTNUM'] + columns_to_merge]

    # Merge the data
    merged_data = pd.merge(target_data, lookup_data_filtered, on='PARTNUM', how='left')

    # Save the sheets back to the target file in the desired order
    with pd.ExcelWriter(target_file, engine='xlsxwriter') as writer:
        original_data.to_excel(writer, sheet_name='Original Data', index=False)
        merged_data.to_excel(writer, sheet_name='Dupes Removed', index=False)
        removed_data.to_excel(writer, sheet_name='Lost Items', index=False)

        # Add formatting to sheets
        workbook = writer.book

        # Define format for wrapped text
        wrap_format = workbook.add_format({'text_wrap': True})

        # Define a color format for the columns being brought in
        color_format = workbook.add_format({'bg_color': '#FFD7E4'})  # Light pink background color

        for sheet_name in ["Original Data", "Dupes Removed"]:
            worksheet = writer.sheets[sheet_name]

            # Freeze panes at column J and just below the first row
            worksheet.freeze_panes(1, 10)

            # Apply the wrapped format to the header row
            for col_num, value in enumerate(original_data.columns.values):
                worksheet.write(0, col_num, value, wrap_format)

                # If the column is one of those being brought in, apply the color format
                if value in ['PSoft Part', 'PSID CT', 'Quoted Mfg', 'Quoted Part', 'Part Class']:
                    worksheet.set_column(col_num, col_num, cell_format=color_format)

            # Enable autofilter for the entire range of data
            worksheet.autofilter(0, 0, len(original_data), len(original_data.columns) - 1)

    messagebox.showinfo("Success", "VLOOKUP operation completed successfully!")
