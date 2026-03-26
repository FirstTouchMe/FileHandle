import os
import eel
import logging
import tkinter as tk
from natsort import natsorted
from logging.handlers import TimedRotatingFileHandler

# 导入你拆分出去的自定义模块
from utils.tk_utils import get_folder_path, get_save_path
from functions.merger_word import core_merge_logic
from functions.merge_txt import merge_txt_files
from functions.converter import DocConverter

# ================= 1. 日志系统配置 =================
LOG_FILE = "word_merger_log.txt"
logger = logging.getLogger("WordMergerLogger")
logger.setLevel(logging.INFO)

if not logger.handlers:
    file_handler = TimedRotatingFileHandler(filename=LOG_FILE, when="D", interval=1, backupCount=3, encoding="utf-8")
    formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

logger.info("========================================")
logger.info("程序启动：模块化 Word 助手已就绪")

# 初始化 Eel (指向你 Vue 编译后的目录)
eel.init('dist')


# ================= 2. Eel 交互接口 =================

@eel.expose
def py_choose_and_scan():
    """选择文件夹并扫描 Word 文件"""
    folder_path = get_folder_path()
    if not folder_path:
        logger.info("用户取消了文件夹选择。")
        return None

    logger.info(f"开始扫描: {folder_path}")
    temp_files = []
    for root_dir, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.docx') and not file.startswith('~$'):
                temp_files.append(os.path.join(root_dir, file))

    sorted_files = natsorted(temp_files)
    return {"folder_path": folder_path, "files": sorted_files}


@eel.expose
def py_choose_save_path(default_name="合并文档.docx", is_pdf=False):
    """根据类型弹出保存框"""
    file_types = [("PDF Document", "*.pdf")] if is_pdf else [("Word Document", "*.docx")]
    return get_save_path(default_name, file_types)


@eel.expose
def py_merge_files(selected_items, save_path):
    """调用 Word 合并逻辑"""

    def update_ui(percent, msg):
        eel.update_progress(percent, msg)

    return core_merge_logic(selected_items, save_path, update_ui)


@eel.expose
def py_merge_txt_files(selected_items, save_path):
    """调用 TXT 合并逻辑"""

    def update_ui(percent, msg):
        eel.update_progress(percent, msg)

    return merge_txt_files(selected_items, save_path, update_ui)


@eel.expose
def py_fast_convert(input_path, mode):
    """
    一键转换接口
    mode: 'to_pdf' 或 'to_word'
    """
    logger.info(f"开始转换任务: {input_path} -> {mode}")
    if mode == 'to_pdf':
        success, result = DocConverter.word_to_pdf(input_path)
    else:
        success, result = DocConverter.pdf_to_word(input_path)

    return {"success": success, "path": result}


# ================= 3. 启动程序 =================

if __name__ == '__main__':
    try:
        # 关闭时 port=0 会自动寻找可用端口
        eel.start('index.html', size=(800, 850), port=0)
    except (SystemExit, KeyboardInterrupt):
        logger.info("程序已正常关闭。")
    except Exception as e:
        logger.critical(f"系统崩溃: {str(e)}", exc_info=True)