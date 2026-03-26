import os
import win32com.client
from pdf2docx import Converter


class DocConverter:
    @staticmethod
    def word_to_pdf(input_docx):
        """利用 Word 原生组件将 docx 转为 pdf"""
        try:
            abs_input = os.path.abspath(input_docx)
            abs_output = os.path.splitext(abs_input)[0] + ".pdf"

            # 初始化 Word 对象
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(abs_input)

            # 17 代表 wdExportFormatPDF
            doc.ExportAsFixedFormat(abs_output, 17)
            doc.Close()
            word.Quit()
            return True, abs_output
        except Exception as e:
            return False, str(e)

    @staticmethod
    def pdf_to_word(input_pdf):
        """将 pdf 转为 docx"""
        try:
            abs_input = os.path.abspath(input_pdf)
            abs_output = os.path.splitext(abs_input)[0] + ".docx"

            cv = Converter(abs_input)
            cv.convert(abs_output)
            cv.close()
            return True, abs_output
        except Exception as e:
            return False, str(e)