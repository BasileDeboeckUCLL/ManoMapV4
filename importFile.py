import pandas as pd
from tkinter import filedialog


def select_input_file(root, label, button_export):
    # Open file dialog and ask user to select an Excel file
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx;*.xls")],
        title="Select an Excel File"
    )
    
    # Check if a file was selected
    if not file_path:
        print("No file selected.")
        label.configure(text="No file selected")
        button_export.configure(state='disabled')
        return None, None

    # Read the Excel file into a DataFrame
    try:
        df = pd.read_excel(file_path)
        print(f"File {file_path} read successfully.")
        # Display the selected file name
        label.configure(text=f"Selected File: {file_path.split('/')[-1]}")
        file_name = file_path.split('/')[-1]
        button_export.configure(state='normal')
        return df, file_name
    except Exception as e:
        print(f"Error reading file: {e}")
        return None, None
