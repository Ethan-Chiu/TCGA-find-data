#---------------------------------------------------------------------------------------
import os.path
import openpyxl
from os import listdir
from os.path import isfile, isdir, join
from pathlib import Path

#---------------------------------------------------------------------------------------
#input
print("Drag in the file: result excel file:")
result_file_name = os.path.basename(input()).rstrip()

print("Drag in the directory: downloaded dir")
path = os.path.basename(input()).rstrip()
#---------------------------------------------------------------------------------------
#all workbook and worksheet
#4 parts need to be changed (result_file_name = 最終資料, wbmata = matadata, wtfwb = 需要找的miR, cliwb = clinical)

wb = openpyxl.load_workbook(result_file_name)


sheet = wb.active
#---------------------------------------------------------------------------------------
#print(Fname_t_Cid)
# 指定要列出所有檔案的目錄

# 取得所有檔案與子目錄名稱
files = listdir(path)
row = 1
for f in files:
  # 產生檔案的完整路徑
  fullpath = join(path, f)
  if isdir(fullpath):
    filefullpath = listdir(fullpath)
    for filename in filefullpath:

      if filename.find("annotations")<0:
        sheet.cell(column = 1, row = row, value = filename)
        row += 1
wb.save(result_file_name)