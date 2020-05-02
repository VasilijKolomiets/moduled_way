#!/usr/bin/env python
# coding: utf-8

# In[ ]:
import pandas as pd
from pathlib import Path
from package_processing import files_reading, package_received
from data_processing import data_processing, excel_writer
from excel_formatting import excel_file_formatting


# Решаем, чтоб в таблица выводились ВСЕ КОЛОНКИ: None -> No Restrictions
pd.options.display.max_columns = None
FOLDER_PATH = Path("E:\\OneDrive\\Проекти\\Chud_Amaz\\Data\\Py\\")

# In[ ]:
#
#  files receiving, reading, pd-frames creations
#
# client_folders = package_receiving()

client_folder = "yumtee"

df_rec, df_adj, df_rei = files_reading(FOLDER_PATH, client_folder)
df_rec, table = data_processing(df_rec, df_adj, df_rei)

print("===ExcelWriter==")
new_file_path = excel_writer(FOLDER_PATH, client_folder, df_rec, table)

print("===Excelformattig==")
excel_file_formatting(str(new_file_path))
