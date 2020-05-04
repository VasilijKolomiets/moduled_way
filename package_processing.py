# -*- coding: utf-8 -*-
"""
Created on Thu Apr 30 14:50:18 2020

@author: Vasil
"""
import pandas as pd
import imaplib
import email
import os
import datetime as dt
from pathlib import Path


# In[ ]
R_A_R_ = ["rec", "adj", "rei"]
imap_url = 'imap.gmail.com'
user = 'vasilij.kolomiets@gmail.com'
password = os.environ['ADMIN_GMAIL_PASSWORD']
'''
imap_url = 'outlook.office365.com' 
user = 'vasilij.kolomiets@outlook.com' 
password ="father1234567890123456789123456789"
'''
imap = imaplib.IMAP4_SSL(imap_url)

# In[]
# from:
# https://myaccount.google.com/lesssecureapps
# https://www.youtube.com/watch?v=bbPwv0TP2UQ
# https://stackoverflow.com/questions/36780993/python-imaplib-search-with-multiple-criteria

'''
    def __init__(self, user, password):
        self.user = user
        self.password = password
        if IMAP_USE_SSL:
            self.imap = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        else:
            self.imap = imaplib.IMAP4(IMAP_SERVER, IMAP_PORT)

    def __enter__(self):
        self.imap.login(self.user, self.password)
        return self

'''


def get_body(msg):
    """

 extracts the body from the email
    Parameters
    ----------
    msg : TYPE
        DESCRIPTION.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None, True)


def searched_letters_uids(search_param_str, imap):
    # search for a particular email
    is_ok, data = imap.uid("search", None, search_param_str)
    assert is_ok == "OK", f"Cant search {search_param_str}"
    return data[0].split()


def suffix_from_now():
    return dt.datetime.now().strftime("_%y%m%d_%H%M%S")


def get_attachments(email_message, folders_path):
    """

    """
    from zipfile import ZipFile

    folder_to_write = ""

    # decoding and glueing the real Subject (if not-ANSI server)
    subject = email.header.decode_header(email_message.get('Subject'))

    def glue(decoded_header: tuple):
        header, coding = decoded_header[0], decoded_header[1]
        return header.decode(coding) if coding else header

    subject = " ".join(map(glue, subject))
    print(f"subject_decoded:: {subject}")

    # get Client name after string "Fwd:" if exists
    if subject:  # Re: Fwd:
        folder_to_write = (subject[4:] if subject[:4] == "Fwd:" else
                           subject).strip().split()[0]

    # make new dir for clients files
    if not folder_to_write:
        folder_to_write = "_" + suffix_from_now()   # add date with secs
    folder_to_write = folders_path / folder_to_write
    if folder_to_write.exists():
        folder_to_write = folder_to_write.with_name(folder_to_write.name + suffix_from_now()) 
    folder_to_write.mkdir()

    # saving attachements to the folder
    for part in email_message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        print(part.get_content_type())
        file_name = part.get_filename()
        print(f"fileName:: {file_name}")
        if bool(file_name):
            with open(folder_to_write / file_name, 'wb') as f:
                f.write(part.get_payload(decode=True))
                #TODO: files_counter  ==1 ?

    with ZipFile(folder_to_write / file_name) as zip_file:
        zip_file.extractall(folder_to_write)
        #TODO: files_counter  >3 ?
    (folder_to_write / file_name).unlink()    # zip_file deleting

    return folder_to_write.name


def get_raw_emails(uids_list):  # extracts emails from byte array
    raw_msgs = []
    for uid in uids_list:
        is_ok, data = imap.uid('fetch', uid, '(RFC822)')
        assert is_ok == "OK", f"Cant fetch {uid}"
        raw_msgs.append(data)
    return raw_msgs


def imap_in(user, password, email_folder="INBOX"):
    imap.login(user, password)
    imap.select(email_folder)


def imap_out():
    imap.close()
    imap.logout()


def package_received(folders_path,
                     sender='ychudnovs@gmail.com',
                     email_folder="INBOX"):

    imap_in(user, password, email_folder)

    # ------  'ychudnovs@gmail.com'  'vikolo@i.ua'
    print(f'\n\n-----{sender}-------\n\n')
    new_folders_with_work = []
    picked_uids = searched_letters_uids(f'(FROM "{sender}" UNSEEN)', imap)
    raw_msgs = get_raw_emails(picked_uids)
    for i, raw_msg in enumerate(raw_msgs):
        email_message = email.message_from_bytes(raw_msg[0][1])
        new_folder_with_work = get_attachments(email_message, folders_path)
        new_folders_with_work.append(new_folder_with_work)
        print(f"{i}:: {get_body(email_message)}\n")
    # ----- close() >>>>  ?????
    imap_out()
    return new_folders_with_work  # organization_dirs


def send_answer(uid, file):
    pass

# In[ ]:
def rename_df_columns(df_):
    '''
    the df columns inplace renaming by replacing character "-" with "_"
    for avoid syntax problems
    '''
    dict_ = {x: x.replace("-", "_") for x in df_.columns}
    df_.rename(columns=dict_, inplace=True)


def file_names_reader(folders_path: Path,
                      client_folder: str,
                      separators: str = " _",
                      patterns: list = R_A_R_) -> dict:
    import re
    """
    files searching in specified directory
    - client_folder-
    with prefix
    - folders_path -
    according patterns in parameter
    - patterns -
    """
    wrk_dir = folders_path / client_folder
    p_compiled = {p: re.compile(f'.*[{separators}]{p}.*\.csv') for p in patterns}
    files = dict()
    for kp, pc in p_compiled.items():
        files[kp] = [f for f in wrk_dir.iterdir()
                     if f.is_file() and pc.search(f.name.lower())]
        if len(files[kp]) > 1:
            raise Exception(f'несколько файлов с подстрокой {kp}')
        if len(files[kp]) == 0:
            raise Exception(f'неn файлов с подстрокой {kp}')

    return {key: file_names[0] for key, file_names in files.items()}


def files_reading(folders_path, client_folder):

    separator = ","  # source file field separators: '\t' = tab or "," = comma

    #  !chardetect direct_adj_20285364484018375.csv
    en_codings = ["cp1250", "cp1252", "cp1253", ]

    files = file_names_reader(folders_path, client_folder)
    pd_files = dict.fromkeys(files.keys())
    readed = False
    for en_coding in en_codings:
        if readed:
            break
        try:
            for key, file_name in files.items():
                pd_files[key] = pd.read_csv(str(file_name),
                                            error_bad_lines=False,
                                            sep=separator,
                                            encoding=en_coding,
                                            )
                rename_df_columns(pd_files[key])
            readed = True
        except ValueError as error:
            print(error)

    if not readed:
        raise Exception(f'неизвестная кодировка. не из {en_codings}')

    return pd_files["rec"], pd_files["adj"], pd_files["rei"]
