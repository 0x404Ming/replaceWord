import os
import subprocess

def open_word_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(".docx"):
                file_path = os.path.join(root, file)
                try:
                    # 使用操作系统命令以 Microsoft Office 打开 Word 文档
                    subprocess.run(["start", "", "/b", "winword", file_path], shell=True)
                    print(f"Opened file: {file_path}")
                except Exception as e:
                    print(f"Error opening file '{file_path}': {e}")

# 指定目录路径
directory_path = r"C:\Users\Q\Desktop\综合办公平台"

# 递归打开目录及其子目录下的 Word 文档
open_word_files(directory_path)
