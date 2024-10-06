import os
import shutil
import sys
import zipfile
import customtkinter as ctk
from tkinter import filedialog, messagebox

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def get_desktop_path():
    """Finding the Desktop Path. The Desktop Path will be different if One Drive or any other
    windows cloud service is installed"""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.exists(desktop_path):
        home_dir = os.path.join(os.path.expanduser("~"), "OneDrive")
        desktop_path = home_dir + "\\Desktop"
    return desktop_path

def extract_files(zip_path):
    """Extracting and sorting the selected zip file as required"""
    if zip_path:
        # Get just the file name
        base_zip_file_name = os.path.basename(zip_path)
        zip_file_name, _ = os.path.splitext(base_zip_file_name)
        desktop_path = get_desktop_path()
        output_dir = os.path.join(desktop_path, zip_file_name)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
    if output_dir:
        file_dict = {}
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            for file_info in zip_ref.infolist():
                file_info = file_info.filename
                if not file_info.endswith('/'):
                    # Extract each file to the output directory
                    extracted_path = zip_ref.extract(file_info, output_dir)
                    # Move the file to the output directory directly if nested
                    final_path = os.path.join(output_dir, os.path.basename(extracted_path))
                    if os.path.exists(final_path):
                        os.remove(final_path)  # Remove if it exists to avoid overwriting
                    os.rename(extracted_path, final_path)
        for item in os.listdir(output_dir):
            item_path = os.path.join(output_dir, item)
            if os.path.isdir(item_path):
                shutil.rmtree(item_path)  # Remove the directory and its contents
            else:
                item_name = os.path.basename(item_path)
                item_name, _ = os.path.splitext(item_name)
                if item_name not in file_dict:
                    file_dict.update({item_name: []})
                else:
                    if item not in file_dict[item_name]:
                        file_dict[item_name].append(item)
                if item not in file_dict[item_name]:
                    file_dict[item_name].append(item)
        dest_sub_dir_name = 1
        for item in file_dict:
            new_path = os.path.join(output_dir, str(dest_sub_dir_name))
            if not os.path.exists(new_path):
                os.makedirs(new_path)
                dest_sub_dir_name += 1
            for file in file_dict[item]:
                source_path = os.path.join(output_dir, file)
                shutil.copy2(source_path, new_path)
                os.remove(source_path)
        messagebox.showinfo("Extraction Complete", "All files have been extracted.")

def merge_files_into_dir(directory_path):
    """Merging files and compressing to zip file"""
    dir_name = os.path.basename(directory_path)
    desktop_path = get_desktop_path()
    doc_dir = os.path.join(desktop_path, 'temp')
    if not os.path.exists(doc_dir):
        os.makedirs(doc_dir)
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith('.docx'):
                source_path = os.path.join(root, file)
                shutil.copy2(source_path, doc_dir)
    zip_directory(doc_dir, dir_name)
    messagebox.showinfo("Merging Completed", "All files have been merged and zipped.")

def zip_directory(directory_path, zip_name):
    """Compressing the specified directory to zip file"""
    desktop_path = get_desktop_path()
    zip_name = 'Abhi_' + zip_name + '.zip'
    zip_file_path = os.path.join(desktop_path, zip_name)
    with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(directory_path):
            for file in files:
                file_path = os.path.join(root, file)
                relative_path = os.path.relpath(file_path, directory_path)
                zipf.write(file_path, relative_path)
    shutil.rmtree(directory_path)


def sort_zip_file():
    file_path = filedialog.askopenfilename(
        title="Select a ZIP file",
        filetypes=[("ZIP files", "*.zip"), ("RAR files", "*.rar")],
    )
    if file_path:
        messagebox.showinfo("File Selected", f"Sorting ZIP file: {file_path}")
        extract_files(file_path)

def merge_directory():
    directory_path = filedialog.askdirectory(title="Select a Directory")
    if directory_path:
        messagebox.showinfo("Directory Selected", f"Merging files from: {directory_path}")
        merge_files_into_dir(directory_path)

# Set up the main application window
ctk.set_appearance_mode("system")  # Options: "light", "dark", "system"
ctk.set_default_color_theme("blue")  # Change theme color

root = ctk.CTk()
root.title("File Organizer")
# Setting icon path for packing in pyinstaller
icon_path = resource_path("resources/img/icon.ico")
root.iconbitmap(icon_path)
root.geometry("300x150")  # Set the window size

# Create buttons with proper alignment
frame = ctk.CTkFrame(root)
frame.pack(pady=20, padx=20, fill="both", expand=True)

btn_sort = ctk.CTkButton(frame, text="Sort ZIP File", command=sort_zip_file)
btn_sort.pack(pady=10)

btn_merge = ctk.CTkButton(frame, text="Merge Directory", command=merge_directory)
btn_merge.pack(pady=10)

# Start the GUI event loop
root.mainloop()
