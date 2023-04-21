import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


class FileRenamerGUI:
    def __init__(self, master):
        self.master = master
        master.title("File Renamer")

        # Create the file selection button and label
        self.select_file_button = tk.Button(master, text="Select Excel File", command=self.select_excel_file)
        self.select_file_button.pack()
        self.selected_file_label = tk.Label(master, text="")
        self.selected_file_label.pack()

        # Create the directory selection button and label
        self.select_dir_button = tk.Button(master, text="Select Directory", command=self.select_directory)
        self.select_dir_button.pack()
        self.selected_dir_label = tk.Label(master, text="")
        self.selected_dir_label.pack()

        # Create the rename files button
        self.rename_files_button = tk.Button(master, text="Rename Files", command=self.rename_files)
        self.rename_files_button.pack()

        # Watch the directory for changes
        self.observer = Observer()
        self.handler = FileRenamerHandler()

    def select_excel_file(self):
        # Open a file dialog to select the Excel file
        excel_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        self.selected_file_label.config(text="Selected Excel File: " + excel_file_path)
        self.excel_file_path = excel_file_path

    def select_directory(self):
        # Open a file dialog to select the directory containing the files to rename
        dir_path = filedialog.askdirectory()
        self.selected_dir_label.config(text="Selected Directory: " + dir_path)
        self.dir_path = dir_path

        # Start watching the directory for changes
        self.handler.set_dir_path(dir_path)
        self.observer.schedule(self.handler, dir_path)
        self.observer.start()

    def rename_files(self):
        # Load the Excel sheet into a pandas DataFrame
        data = pd.read_excel(self.excel_file_path)

        # Iterate over each row in the DataFrame
        for index, row in data.iterrows():
            # Get the old and new file names from the row
            old_file_name = os.path.join(self.dir_path, row['Old File Name'])
            new_file_name = os.path.join(self.dir_path, row['New File Name'])

            # Rename the file
            os.rename(old_file_name, new_file_name)

            # Update the Excel sheet with the new file name
            data.at[index, 'New File Name'] = os.path.basename(new_file_name)
            data.to_excel(self.excel_file_path, index=False)

        # Show a message box when the files have been renamed
        messagebox.showinfo("File Renamer", "Files have been renamed!")

    def stop_watching(self):
        # Stop watching the directory for changes
        self.observer.stop()
        self.observer.join()


class FileRenamerHandler(FileSystemEventHandler):
    def set_dir_path(self, dir_path):
        self.dir_path = dir_path

    def on_modified(self, event):
        # Check if the modified file is an Excel file in the selected directory
        if event.is_directory or not event.src_path.endswith('.xlsx') or not event.src_path.startswith(
                self.dir_path):
            return

        # Load the Excel sheet into a pandas DataFrame
        data = pd.read_excel(self.excel_file_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = FileRenamerGUI(root)
    root.mainloop()
    app.stop_watching()