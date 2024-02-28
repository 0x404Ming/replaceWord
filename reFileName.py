import os

def rename_files_and_dirs(directory, replacements):
    # First, rename files.
    for root, dirs, files in os.walk(directory):
        for file in files:
            # Skip temporary files that start with '~$'
            if file.startswith("~$"):
                continue

            new_file_name = file
            for old_str, new_str in replacements:
                new_file_name = new_file_name.replace(old_str, new_str)
            
            if new_file_name != file:
                file_path = os.path.join(root, file)
                new_file_path = os.path.join(root, new_file_name)
                try:
                    os.rename(file_path, new_file_path)
                    print(f"Renamed file: {file_path} -> {new_file_path}")
                except Exception as e:
                    print(f"Error renaming file '{file_path}': {e}")
    
    # Next, rename directories.
    for root, dirs, _ in os.walk(directory, topdown=False):  # topdown=False is necessary to handle nested dirs
        for dir in dirs:
            new_dir_name = dir
            for old_str, new_str in replacements:
                new_dir_name = new_dir_name.replace(old_str, new_str)
            
            if new_dir_name != dir:
                dir_path = os.path.join(root, dir)
                new_dir_path = os.path.join(root, new_dir_name)
                try:
                    os.rename(dir_path, new_dir_path)
                    print(f"Renamed directory: {dir_path} -> {new_dir_path}")
                except Exception as e:
                    print(f"Error renaming directory '{dir_path}': {e}")

# The directory path
directory_path = r"C:\Users\Q\Desktop\过程文档V"

# List of old strings and corresponding new strings
replacements = [
    ("002121", "aa"),
    ("old_str","new_str")
        # 添加更多的替换对，如：("old_str2", "new_str2")
]

rename_files_and_dirs(directory_path, replacements)
