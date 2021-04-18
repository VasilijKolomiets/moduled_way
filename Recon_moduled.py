"""Yes."""
import pandas as pd
from pathlib import Path
import package_processing as pp
from data_processing import data_processing, excel_writer
from excel_formatting import excel_file_formatting, current_os_is_win

#import scheduleo
import time

"""
schedule.every(10).seconds.do(job)
schedule.every(10).minutes.do(job)
schedule.every().hour.do(job)
schedule.every().day.at("10:30").do(job)
schedule.every(5).to(10).minutes.do(job)
schedule.every().monday.do(job)
schedule.every().wednesday.at("13:15").do(job)
schedule.every().minute.at(":17").do(job)

while True:
    schedule.run_pending()
    time.sleep(1)
"""

# Решаем, чтоб в таблица выводились ВСЕ КОЛОНКИ: None -> No Restrictions
pd.options.display.max_columns = None
FOLDERS_PATH = Path(r"D:\_\Snapshot")
JOCK_FILE = Path(r"D:\OneDrive\PyCodes\SHEDULER\tenor_Mister_Bin.gif")

# In[ ]:
#
# #  sender='vikolo@i.ua'
#

boss_emails = ['vikolo@i.ua', 'Yu <ychudnovs@gmail.com>', ]  #


def main():

    import datetime

    with open("senders.txt", "r") as senders_file:
        senders = [sender.strip() for sender in senders_file.read().splitlines()] 

    for sender in senders:
        for uid in pp.packadges_uids(sender=sender):

            print(f"\n\n new iter with uid = {uid}\n\n")
            client_folder_name, email_message = pp.package_received(FOLDERS_PATH, uid)
            if any(map(client_folder_name.count, ['not ZIP file attached',
                                                  'not three files in ZIP',
                                                  'brocken *.zip'])
                   ):
                print(client_folder_name)
                continue
        # %%  good three files in ZIP file attached
            df_rec, df_adj, df_rei = pp.files_reading(FOLDERS_PATH, client_folder_name)

            print("===data_processing==")
            files = data_processing(df_rec, df_adj, df_rei)
            # check if any DF is empty
            if all(tuple(map(lambda d_f: not d_f.empty, files.values()))):
                print("===ExcelWriter==")
                client_name = client_folder_name
                new_files_path = excel_writer(FOLDERS_PATH, client_folder_name, client_name, files)

                print("===Excelformattig==")
                excel_file_formatting(str(new_files_path["xlsx"]))
                body_text = "Macros reply - Ok"
            else:
                new_files_path = {"xlsx": JOCK_FILE}
                body_text = "Nothing to do"

            # %%
            print("===send_answer==")
            pp.send_answer(email_message, new_files_path["xlsx"], body_text=body_text,
                           cc_addrs=boss_emails)
            if current_os_is_win:
                folder = FOLDERS_PATH / Path(client_folder_name)
                [f.unlink() for f in folder.glob("*") if f.is_file()]
                folder.rmdir()
        pp.imap_out()
    print(f"One time done {datetime.datetime.now()}")
    return None


if __name__ == "__main__":
    '''
    print(f"start af {time.time()}")
    schedule.every(2).hours.do(main)
    schedule.run_all()

    while True:
        schedule.run_pending()
        time.sleep(120)
        print(".", end="")

    '''

    while True:
        main()
        time.sleep(60*60*2)
