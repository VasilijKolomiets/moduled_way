"""Yes."""
import set_path

import pandas as pd
from pathlib import Path
import package_processing as pp
from data_processing import (
    files_reading, data_processing, excel_writer,
    fee_files_reading, fee_data_processing, fee_excel_writer,
)

from excel_formatting import (
    current_os_is_win,
    excel_file_formatting, fee_excel_file_formatting
)

# import scheduleo
import time

dir(set_path)


# Решаем, чтоб в таблица выводились ВСЕ КОЛОНКИ: None -> No Restrictions
pd.options.display.max_columns = None
FOLDERS_PATH = Path(r"D:\_")
JOCK_FILE = Path(r"D:\OD\OneDrive\PyCodes\SHEDULER\tenor_Mister_Bin.gif")

# In[ ]:
#

# 'vikolo@i.ua'  'Yu <ychudnovs@gmail.com>'
boss_emails = ['vikolo@i.ua', 'Yu <ychudnovs@gmail.com>']


def main():
    """Calculate Amazon reports."""
    import datetime

    with open("senders.txt", "r") as senders_file:    # list of registered file senders
        senders = [sender.strip() for sender in senders_file.read().splitlines()]

    for sender in senders:
        for uid in pp.packadges_uids(sender=sender):
            print(f"\n\n new iter with uid = {uid}\n\n")
            email_message = pp.get_email(uid)
            # TODO:
            received = pp.package_received(FOLDERS_PATH, email_message)
            if not received["exit_is_ok"]:
                print(received["folder"], "--->", received["exit_message"])
                # TODO: send message
                continue    # TODO:  corect loop exiting

            task_name = received["task_type"]
            path_to_files_folder = received["folder"]
            if task_name == "Snapshots":
                #  good three files in ZIP file attached
                readed = files_reading(path_to_files_folder)
            elif task_name == "FEE":
                readed = fee_files_reading(path_to_files_folder)
            else:
                assert False, "WHAT?"

            pd_files = readed["files"]

            if not readed["exit_is_ok"]:
                print(received["folder"], "--->", readed["exit_message"])
                # TODO: send message
                continue    # TODO:  corect loop exiting

            print("===data_processing==")
            if task_name == "Snapshots":
                data_response = data_processing(**pd_files, OB=None)
            elif task_name == "FEE":    # TODO:
                data_response = fee_data_processing(**pd_files)
            else:
                assert False, "WHAT?"

            if not data_response["exit_is_ok"]:
                print(received["folder"], "--->", data_response["exit_message"])
                # TODO: send message
                continue    # TODO:  correct loop exiting

            files = data_response["files"]
            client_name = path_to_files_folder.name
            if task_name == "Snapshots":
                if all(tuple(map(lambda d_f: not d_f.empty, files.values()))):
                    print("===ExcelWriter==")

                    new_files_path = excel_writer(path_to_files_folder, client_name, files)

                    print("===Excelformattig==")
                    excel_file_formatting(str(new_files_path["xlsx"]))
                    body_text = "Macros reply - Ok"
                else:
                    new_files_path = {"xlsx": JOCK_FILE}
                    body_text = "Nothing to do"
                    # TODO: send message  ???

            elif task_name == "FEE":
                xl_response = fee_excel_writer(path_to_files_folder, client_name, files)
                if not xl_response["exit_is_ok"]:
                    print(received["folder"], "--->", xl_response["exit_message"])
                    # TODO: send message
                    continue    # TODO:  correct loop exiting

                format_xl_response = fee_excel_file_formatting(xl_response["xlsx_files"])
                if not format_xl_response["exit_is_ok"]:
                    print(received["folder"], "--->", format_xl_response["exit_message"])
                    # TODO: send message
                    continue    # TODO:  correct loop exiting

                zip_to_send = format_xl_response['zip_path']
                body_text = "Macros reply - Ok"
            else:
                assert False, "WHAT?"
            # check if any DF is empty

            # %%
            print("===send_answer==")
            if task_name == "Snapshots":
                pp.send_answer(email_message,
                               new_files_path["xlsx"],
                               body_text=body_text,
                               cc_addrs=boss_emails)
            elif task_name == "FEE":
                pp.send_answer(email_message,
                               zip_to_send,
                               body_text=body_text,
                               cc_addrs=boss_emails)
            else:
                assert False, "WHAT?"

            if not current_os_is_win:
                [f.unlink() for f in path_to_files_folder.glob("*") if f.is_file()]
                path_to_files_folder.rmdir()

        pp.imap_out()
    print(f"One time done {datetime.datetime.now()}")
    return None


if __name__ == "__main__":

    while True:
        main()
        time.sleep(20)


'''
    print(f"start af {time.time()}")
    schedule.every(2).hours.do(main)
    schedule.run_all()

    while True:
        schedule.run_pending()
        time.sleep(120)
        print(".", end="")

'''
