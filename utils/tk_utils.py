import tkinter as tk
from tkinter import filedialog
import logging

logger = logging.getLogger("WordMergerLogger")


def get_folder_path():
    """弹出选择文件夹对话框"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory(parent=root, title="选择存放 Word 文件的文件夹")
    root.destroy()
    return folder_path


def get_save_path(default_name="合并完成的文档.docx", file_types=None):
    """弹出保存文件对话框"""
    if file_types is None:
        file_types = [("Word Document", "*.docx")]

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    save_path = filedialog.asksaveasfilename(
        parent=root,
        title="保存文件",
        defaultextension=".docx" if "docx" in str(file_types) else ".pdf",
        filetypes=file_types,
        initialfile=default_name
    )
    root.destroy()
    return save_path