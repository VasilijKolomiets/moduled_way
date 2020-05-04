#!/usr/bin/env python
# coding: utf-8

# In[ ]:
import pandas as pd
from pathlib import Path
from package_processing import files_reading, package_received, send_answer
from data_processing import data_processing, excel_writer
from excel_formatting import excel_file_formatting


# Решаем, чтоб в таблица выводились ВСЕ КОЛОНКИ: None -> No Restrictions
pd.options.display.max_columns = None
FOLDERS_PATH = Path("E:\\OneDrive\\Проекти\\Chud_Amaz\\Data\\Py\\")

# In[ ]:
#
#
for client_folder in package_received(FOLDERS_PATH, sender = 'vikolo@i.ua'):  # , sender = 'vikolo@i.ua'
    df_rec, df_adj, df_rei = files_reading(FOLDERS_PATH, client_folder)
    df_rec, table = data_processing(df_rec, df_adj, df_rei)

    print("===ExcelWriter==")
    new_file_path = excel_writer(FOLDERS_PATH, client_folder, df_rec, table)

    print("===Excelformattig==")
    excel_file_formatting(str(new_file_path))

