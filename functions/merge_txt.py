# functions/merge_txt.py - 新建 TXT 合并模块
import os
import logging

logger = logging.getLogger("WordMergerLogger")


def merge_txt_files(selected_items, save_path, progress_callback):
    """
    合并 TXT 文件，将文件名作为目录标题
    :param selected_items: 文件路径列表
    :param save_path: 保存路径
    :param progress_callback: 进度回调函数
    """
    total_files = len(selected_items)
    logger.info(f"=== 开始执行 TXT 合并任务，计划合并：{total_files} 个文件 ===")
    
    if total_files == 0:
        return {"status": "error", "msg": "没有需要合并的文件！"}
    
    try:
        with open(save_path, 'w', encoding='utf-8') as outfile:
            for i, file_path in enumerate(selected_items):
                percent = int(((i + 1) / total_files) * 100)
                file_name = os.path.basename(file_path)
                file_name_without_ext = os.path.splitext(file_name)[0]
                
                progress_callback(percent, f"正在处理：{file_name}")
                logger.info(f"正在合并：{file_path}")
                
                outfile.write(f"\n{'='*50}\n")
                outfile.write(f"{file_name_without_ext}\n")
                outfile.write(f"{'='*50}\n\n")
                
                with open(file_path, 'r', encoding='utf-8') as infile:
                    content = infile.read()
                    outfile.write(content)
                    if not content.endswith('\n'):
                        outfile.write('\n')
                
                outfile.write("\n")
        
        progress_callback(100, "合并完成！")
        logger.info(f"=== TXT 合并任务成功完成，保存至：{save_path} ===")
        return {"status": "success", "msg": f"🎉 成功合并了 {total_files} 个 TXT 文件！\n\n已保存至:\n{save_path}"}
    
    except Exception as e:
        logger.error(f"TXT 合并失败：{str(e)}", exc_info=True)
        return {"status": "error", "msg": f"发生错误：{str(e)}"}
