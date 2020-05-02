# -*- coding: utf-8 -*-
"""
Created on Thu Apr 30 14:50:18 2020

@author: Vasil
"""
import pandas as pd
import imaplib
import email
import os
from pathlib import Path

# In[ ]
R_A_R_ = ["rec", "adj", "rei"]
imap_url = 'imap.gmail.com'
user = 'vasilij.kolomiets@gmail.com'
password = os.environ['ADMIN_GMAIL_PASSWORD']
attachment_dir = '..\\\\Data\\Py\\Saved\\'


# In[]
imap = imaplib.IMAP4_SSL(imap_url)
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
# search for a particular email
def searched_letters_uids(search_param_str, imap):
    is_ok, data = imap.uid("search", None, search_param_str)
    assert is_ok == "OK", f"Cant search {search_param_str}"
    return data[0].split()


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


def get_attachments(msg):
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        
        print(part.get_content_type()) 
        fileName = part.get_filename()
#        try:
#            if part.
        print(f"fileName:: {fileName}")
        if bool(fileName):
            subject = part.get('Subject')
            folder = ""
            if subject:
                folder = (subject[4:] if subject[:4]=="Fwd:" 
                          else subject).strip().split()[0]
            if not folder:
                folder = "_"
            filePath = os.path.join(attachment_dir, folder, fileName)
            ##  os.mdir(filepeth) %%% 
            with open(filePath, 'wb') as f:
                f.write(part.get_payload(decode=True))


# extracts emails from byte array
def get_emails(result_bytes):
    msgs = []
    for num in result_bytes[0].split():
        typ, data = imap.fetch(num, '(RFC822)')
        msgs.append(data)
    return msgs


def package_received():
    from zipfile import ZipFile
    imap.login(user, password)
    imap.list() 
    imap.select("INBOX")
    
    # ------
    #    'ychudnovs@gmail.com'  'vikolo@i.ua'
    sender = 'vikolo@i.ua'
    print(f'\n\n-----{sender}-------\n\n')
    msgs = get_emails(search("FROM", sender, imap))
    for i, msg in enumerate(msgs):
        raw_ = email.message_from_bytes(msg[0][1]) 
        get_attachments(raw_)
        print(f"{i}:: {get_body(raw_)}\n")
        # ----- close() >>>>  ?????
    return None  # organization_dir


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
        files[kp] = [f for f in wrk_dir.iterdir() if f.is_file() and pc.search(f.name.lower())]
        if len(files[kp])>1:
            raise Exception(f'несколько файлов с подстрокой {kp}')
        if len(files[kp])==0:
            raise Exception(f'неn файлов с подстрокой {kp}')    
        
    return {key: file_names[0] for key, file_names in files.items()}


def files_reading(folders_path, client_folder):
    
    separator = ","  # source file field separators: '\t' = tab or "," = comma    
    
    #  !chardetect direct_adj_20285364484018375.csv
    en_codings = ["cp1250", "cp1252", "cp1253", ] # "cp1250", "cp1252", "cp1253", 
    
    files = file_names_reader(folders_path, client_folder)
    pd_files = dict.fromkeys(files.keys())
    readed = False
    for en_coding in en_codings:
        if readed: break
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


