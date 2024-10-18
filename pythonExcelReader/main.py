import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk  # Import ttk for the progress bar


# Function to read the Excel file and create folders and files
def create_folder_and_files(path_excel, target_path):
    try:
        # Load the Excel file
        dataframe = pd.read_excel(path_excel, header=None, skiprows=1)
        dataframe.columns = ['Folders', 'Files', 'Content']

        # Get the total number of rows (for the progress bar)
        total_rows = len(dataframe)
        progress_bar['maximum'] = total_rows  # Set the maximum value for the progress bar

        if total_rows == 0:
            messagebox.showwarning("Warning", "The Excel file is empty!")
            return

        # Iterate over the rows of the sheet (starting from row 0)
        for index, row in dataframe.iterrows():

            # Extract data from the columns
            folder_name = os.path.join(target_path, row['Folders'])  # TARGET_PATH + folder name
            file_name = row['Files']  # Column B (2)
            file_content = row['Content']  # Column C (3)

            print(f"[{folder_name}: {file_name} ( {file_content} )]")

            # Create the folder if it doesn't exist
            if folder_name and not os.path.exists(folder_name):
                os.makedirs(folder_name)

            # Full file path
            if file_name:
                file_path = os.path.join(str(folder_name), file_name)

                # Create the file and write the content
                with open(file_path, 'w') as file:
                    if file_content:
                        file.write(file_content)
                    else:
                        file.write('')

            # Update the progress bar
            progress_bar['value'] = index + 1  # Increment the progress bar value
            root.update_idletasks()  # Update the GUI to reflect changes

        messagebox.showinfo("Success", "Folders and files created successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
    finally:
        progress_bar['value'] = 0  # Reset the progress bar at the end


# Function to open the file dialog and select the Excel file
def excel_file_selection():
    path_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry_excel_file.delete(0, tk.END)
    entry_excel_file.insert(0, path_file)


# Function to open the file dialog and select the save directory
def target_path_selection():
    directory = filedialog.askdirectory()
    entry_target_path.delete(0, tk.END)
    entry_target_path.insert(0, directory)


# Function to start the process on button click
def start_creation():
    path_excel = entry_excel_file.get()
    path_directory = entry_target_path.get()

    if path_excel and path_directory:
        create_folder_and_files(path_excel, path_directory)
    else:
        messagebox.showwarning("Warning", "Please enter all required values.")


# Create the main window
root = tk.Tk()
root.title("Create Folders and Files from Excel")

# ----------------------------------[Select Excel file:] [___________________________] [Browse]
# Label and field for the Excel file path
tk.Label(root, text="Select the Excel file:").grid(row=0, column=0, padx=10, pady=10)
entry_excel_file = tk.Entry(root, width=40)
entry_excel_file.grid(row=0, column=1, padx=10, pady=10)
btn_excel_file = tk.Button(root, text="Browse", command=excel_file_selection)
btn_excel_file.grid(row=0, column=2, padx=10, pady=10)

# ----------------------------------[Select save directory:] [___________________________] [Browse]
# Label and field for the save directory
tk.Label(root, text="Select the save directory:").grid(row=1, column=0, padx=10, pady=10)
entry_target_path = tk.Entry(root, width=40)
entry_target_path.grid(row=1, column=1, padx=10, pady=10)
btn_target_path = tk.Button(root, text="Browse", command=target_path_selection)
btn_target_path.grid(row=1, column=2, padx=10, pady=10)

# Button to start the process
start_btn = tk.Button(root, text="Create Folders and Files", command=start_creation)
start_btn.grid(row=2, column=0, columnspan=3, padx=10, pady=20)

# Progress bar
progress_bar = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=300)
progress_bar.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

# Start the main window
root.mainloop()
