# Import necessary libraries
import tkinter as tk
from tkinter import filedialog
import os
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from my_tools import create_mv2, xls_to_dict, TRD_date_range, get_mv2_types


def browse_file():
    """
    Opens a file dialog for the user to select an Excel file.
    Updates the file path entry widget with the selected file's path.
    """
    file_path = filedialog.askopenfilename()
    if file_path:
        file_path_var.set(file_path)


def create():
    """
    Processes the selected Excel file to create a new MV2 file.
    Handles errors and displays the status to the user.
    """
    xl_file_path = file_path_var.get()
    if not xl_file_path:
        status_label.config(
            text="Please select a file first.", bootstyle="danger")
        return

    try:

        # Get TRD date range
        TRD_start_hour, TRD_end_hour = TRD_date_range(xl_file_path)
        print(
            f"TRD start hour: {TRD_start_hour}, TRD end hour: {TRD_end_hour}")

        types = get_mv2_types(xl_file_path) + ["ALL"]
        print(f"MV2 types to process: {types}")

        for mv2_type in types:
            print(f"Processing type: {mv2_type}")
            # Process the Excel file to get the mv2 dictionary
            mv2 = xls_to_dict(xl_file_path, TRD_start_hour,
                              TRD_end_hour, mv2_type)
            print("Data processed successfully for type:", mv2_type)

            # Create a new MV2 Excel file
            directory = os.path.dirname(xl_file_path)
            new_file_path = create_mv2(
                mv2, TRD_start_hour, TRD_end_hour, directory, mv2_type)
            print(f"New MV2 file created at: {new_file_path}")

        status_label.config(
            text=f"Successfully creating MV2 files for {len(types)} types.", bootstyle="success"
        )

    except Exception as e:
        status_label.config(text=f"An error occurred: {e}", bootstyle="danger")


# --- GUI Setup ---
# Create the main application window
root = ttk.Window(themename="morph")
root.title("MV2 Creator App")
root.geometry("700x300")
root.resizable(False, False)

# --- Widgets ---
# Create a frame to hold all other widgets
main_frame = ttk.Frame(root, padding=20)
main_frame.pack(fill=BOTH, expand=True)

# Title Label for the application
title_label = ttk.Label(
    main_frame, text="MV2 Creator App", font=("Arial", 16, "bold"))
title_label.pack(pady=(20, 0))

# Subtitle Label
title_label = ttk.Label(main_frame, text="By anas asimi", font=("Arial", 12))
title_label.pack(pady=(0, 20))

# Create a frame for the file path entry and browse button
file_frame = ttk.Frame(main_frame)
file_frame.pack(pady=5, fill=X)

# Create a variable and an entry widget for the file path
file_path_var = tk.StringVar()
file_entry = ttk.Entry(file_frame, textvariable=file_path_var)
file_entry.pack(side=LEFT, expand=True, fill=X)

# Create a button to browse for a file
browse_button = ttk.Button(
    file_frame, text="Browse", command=browse_file, bootstyle="primary"
)
browse_button.pack(side=RIGHT, padx=(5, 0))

# Create a button to initiate the file creation process
create_button = ttk.Button(
    main_frame, text="Create File", command=create, bootstyle="primary"
)
create_button.pack(pady=10)

# Create a label to display the status of the operation
status_label = ttk.Label(main_frame, text="")
status_label.pack(pady=(0, 5))

# Start the main event loop of the application
root.mainloop()
