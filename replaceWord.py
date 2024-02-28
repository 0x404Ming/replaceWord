import os
import win32com.client as win32
import logging

def setup_logging(directory_path):
    log_filename = os.path.basename(directory_path) + "_replacement.log"
    logging.basicConfig(level=logging.INFO, 
                        format='%(asctime)s - %(levelname)s: %(message)s',
                        handlers=[logging.FileHandler(log_filename, 'w', 'utf-8'), logging.StreamHandler()])

def set_file_write_permission(directory):
    for root, _, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            os.chmod(file_path, 0o777)  # Setting write permission

def replace_string_in_doc(file_path, replacements):
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(file_path)
    word.Visible = False

    total_replacements_made = 0
    for old_str, new_str in replacements:
        rng = doc.Content
        find = rng.Find
        find.ClearFormatting()
        find.Text = old_str
        find.Replacement.Text = new_str
        find.Replacement.ClearFormatting()
        # 使用 wdReplaceAll 替换整个文档中的所有实例
        replacements_made = find.Execute(FindText=old_str, Forward=True, Replace=win32.constants.wdReplaceAll)
        total_replacements_made += replacements_made if replacements_made else 0
        
        # 处理页眉
        for section in doc.Sections:
            header = section.Headers(win32.constants.wdHeaderFooterPrimary).Range
            find = header.Find
            find.ClearFormatting()
            find.Text = old_str
            find.Replacement.Text = new_str
            find.Replacement.ClearFormatting()
            
            header_replacements = find.Execute(FindText=old_str, Forward=True, Replace=win32.constants.wdReplaceAll)
            total_replacements_made += header_replacements if header_replacements else 0

    doc.Save()
    doc.Close()
    word.Quit()

    return total_replacements_made

def process_directory(directory, replacements):
    total_replacements = 0
    for root, _, files in os.walk(directory):
        for file in files:
            if file.startswith("~$"):
                continue  # skip temporary Word files

            file_path = os.path.join(root, file)
            if file.endswith(".doc") or file.endswith(".docx"):
                try:
                    replacements_made = replace_string_in_doc(file_path, replacements)
                    total_replacements += replacements_made
                    logging.info(f"Processed: {file_path}, Replacements Made: {replacements_made}")
                except Exception as e:
                    logging.error(f"Error processing '{file_path}': {e}")

    logging.info(f"Total Replacements Made: {total_replacements}")




if __name__ == "__main__":
    #设置要替换文档所在路径 替换所有在该路径下的word文档内容
    a = r'C:\Users\Q\Desktop\新建文件夹'
    directory_path = a
    setup_logging(directory_path)
    # 一个列表，包含需要替换的old_str和new_str的元组
    replacements = [

        ("old_string","new_string")

        # 添加更多的替换对，如：("old_str2", "new_str2")
    ]
    

    set_file_write_permission(directory_path)
    if os.path.exists(directory_path):
        process_directory(directory_path, replacements)
    else:
        logging.error(f"Directory does not exist: {directory_path}")
