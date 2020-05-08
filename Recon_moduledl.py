#!/usr/bin/env python
# coding: utf-8

# In[ ]:
import pandas as pd
from pathlib import Path
import package_processing as pp
from data_processing import data_processing, excel_writer
from excel_formatting import excel_file_formatting


# Решаем, чтоб в таблица выводились ВСЕ КОЛОНКИ: None -> No Restrictions
pd.options.display.max_columns = None
FOLDERS_PATH = Path(r"D:\_\_Py_")

# In[ ]:
#
# #  sender='vikolo@i.ua'
#
for uid in pp.packadges_uids():

    print(f"\n\n new iter with uid = {uid}\n\n")
    client_folder, email_message = pp.package_received(FOLDERS_PATH, uid)
    df_rec, df_adj, df_rei = pp.files_reading(FOLDERS_PATH, client_folder)

    print("===data_processing==")
    df_rec, table = data_processing(df_rec, df_adj, df_rei)

    print("===ExcelWriter==")
    new_file_path = excel_writer(FOLDERS_PATH, client_folder, df_rec, table)

    print("===Excelformattig==")
    excel_file_formatting(str(new_file_path))

    print("===send_answer==")
    pp.send_answer(email_message, new_file_path)
