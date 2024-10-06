import shutil
import sys
import zipfile
import os

current_working_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
os.chdir(current_working_dir)

user_input = int(input("Press 1 for extracting files\nPress 2 for merging doc files\n"))


def sort_files():
    dir_elements = os.listdir(current_working_dir)
    for el in dir_elements:
        while el.endswith('.zip'):
            zip_name = el.split('.')[0]
            extract_dir_path = os.path.join(current_working_dir, zip_name)
            zip_file_path = os.path.join(current_working_dir, el)
            break
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        for name in zip_ref.namelist():
            if not name.endswith('/'):
                if not os.path.exists(extract_dir_path):
                    os.makedirs(extract_dir_path)
                extract_to = extract_dir_path
            else:
                extract_to = current_working_dir
        zip_ref.extractall(extract_to)
    source_dir = extract_to
    extract_to = extract_to + '/' + zip_name
    # print("Extracting files...\n")
    copy_files(source_dir, extract_to)
    shutil.rmtree(extract_to)

def copy_files(source_dir, extract_dir):
    file_dict = {}
    sorted_dir = os.path.join(source_dir, 'Sorted Docs')
    if not os.path.exists(sorted_dir):
        os.makedirs(sorted_dir)
    sorted_dir_path = sorted_dir
    for item in os.listdir(extract_dir):
        item_name = item.split('.')[0]
        if item_name not in file_dict:
            file_dict.update({item_name: []})
        else:
            if item not in file_dict[item_name]:
                file_dict[item_name].append(item)
        if item not in file_dict[item_name]:
            file_dict[item_name].append(item)
    dest_sub_dir_name = 1
    for item in file_dict:
        new_path = os.path.join(sorted_dir_path, str(dest_sub_dir_name))
        if not os.path.exists(new_path):
            os.makedirs(new_path)
            dest_sub_dir_name += 1
        for file in file_dict[item]:
            source_path = os.path.join(extract_dir, file)
            shutil.copy2(source_path, new_path)
    print("Files sorted successfully!\n")


def sorted_dir():
    dir_elements = os.listdir(current_working_dir)
    for el in dir_elements:
        while el.endswith('.zip'):
            zip_name = el.split('.')[0]
            sorted_dir_path = os.path.join(current_working_dir, zip_name)
            break
    doc_dir = os.path.join(sorted_dir_path, 'Final Docs')
    if not os.path.exists(doc_dir):
        os.makedirs(doc_dir)
    sorted_dir_path = sorted_dir_path + "/" + 'Sorted Docs'
    if os.path.exists(sorted_dir_path):
        for dir in os.listdir(sorted_dir_path):
            doc_loc = sorted_dir_path + "/" + dir
            for file in os.listdir(doc_loc):
                if file.endswith('.docx'):
                    source_path = os.path.join(doc_loc, file)
                    shutil.copy2(source_path, doc_dir)
    zip_directory(doc_dir, zip_name)
    print("Final docs ready!\n")


def zip_directory(directory_path, zip_name):
    print("Zipping..\n")
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.exists(desktop_path):
        home_dir = os.path.join(os.path.expanduser("~"), "OneDrive")
        desktop_path = home_dir + "\\Desktop"
    zip_name = 'Abhi_' + zip_name + '.zip'
    zip_file_path = os.path.join(desktop_path, zip_name)
    with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(directory_path):
            for file in files:
                file_path = os.path.join(root, file)
                relative_path = os.path.relpath(file_path, directory_path)
                zipf.write(file_path, relative_path)
    shutil.rmtree(directory_path)


if user_input == 1:
    sort_files()
if user_input == 2:
    sorted_dir()
input("Press any key to exit")
