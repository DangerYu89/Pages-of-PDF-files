from glob import glob
import os
import xlsxwriter
import time
from tkinter import filedialog
import PyPDF2


def desktop_path():
    return os.path.join(os.path.expanduser('~'), "Desktop")


folder = filedialog.askdirectory(title="选择要统计PDF的文件夹:", initialdir=desktop_path())
pdf_list = glob(os.path.join(folder, "*.pdf"), recursive=True)
pages = []
lt = time.localtime()  # UTC格式当前时区时间
st = time.strftime("%H时%M分%S秒", lt)  # 转换为格式化日期
wb = xlsxwriter.Workbook("%s.xlsx" % st)
ws = wb.add_worksheet("test")
ws.write('A1', '页数')
ws.write('B1', 'pdf文件路径')
ws.write_column(1, 1, pdf_list)
for pdf in pdf_list:
    reader = PyPDF2.PdfFileReader(pdf)
    num_page = reader.getNumPages()
    pages.append(num_page)
ws.write_column(1, 0, pages)
wb.close()
print("已完成！")
print("%s 的pdf文件页数信息已保存为%s.xlsx工作薄！" % (folder, st))
