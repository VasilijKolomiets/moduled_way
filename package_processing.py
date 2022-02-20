"""
Created on Thu Apr 30 14:50:18 2020

@author: Vasil
"""
import pandas as pd
import imaplib
import email
import os
import sys
import datetime as dt
from pathlib import Path

import shutil


# In[ ]
imap_url = 'imap.gmail.com'
user = 'vasilij.kolomiets@gmail.com'
password = "aggclkaseqdgfucg"  # os.environ['ADMIN_GMAIL_PASSWORD']
'''
imap_url = 'outlook.office365.com'
user = 'vasilij.kolomiets@outlook.com'
password ="father1234567890123456789123456789"
'''
# imap = imaplib.IMAP4_SSL(imap_url)
imap = None

# In[ ]


def suffix_from_now():
    """Get Year-month-day-time as a string without separators."""
    return dt.datetime.now().strftime("_%y%m%d_%H%M%S")


def create_folder(folders_path, folder_to_create):
    """Check if exist folder = rename and create filder."""
    folder_to_write = folders_path / folder_to_create
    if folder_to_write.exists():  # add date with secs
        folder_to_write = folder_to_write.with_name(folder_to_write.name + suffix_from_now())
    folder_to_write.mkdir(mode=0o777, parents=True)

    return folder_to_write


def files_moveup_from(folder_path):
    """Remove empty folder after unziping."""
    for file in folder_path.iterdir():
        file.rename(file.parent.parent / file.name)
    folder_path.rmdir()  # removing emptied folder

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
    Extract the body from the email.

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


def searched_emails_uids(search_param_str, imap):
    """Search for a particular email."""
    is_ok, data = imap.uid("search", None, search_param_str)
    assert is_ok == "OK", f"Cant search {search_param_str}"
    return data[0].split()


def get_attachments(email_message, folders_path) -> tuple:
    """Get 'email_message' attachments to the folder 'folders_path'."""
    import shutil      # !!!  not used
    from email.header import decode_header

    folder_to_write = ""

    # decoding and glueing the real Subject (if not-ANSI server)
    subject = decode_header(email_message.get('Subject'))

    def glue(decoded_header: tuple):
        header, coding = decoded_header[0], decoded_header[1]
        return header.decode(coding) if coding else header

    subject = " ".join(map(glue, subject))
    print(f"subject_decoded:: {subject}")

    # get Client name after string "Fwd:" if exists %%%
    subject = subject.replace("Re:", "")
    subject = subject.replace("Fwd:", "")

    folder_to_create = subject.strip().split()[0] if subject else "was_no_subj"

    # make new dir for clients files
    folder_to_write = create_folder(folders_path, folder_to_create)

    # TODO: process situation when email  has no attachement (only link or )
    file_name = 'was_no_attachements'
    # saving attachements to the folder
    for part in email_message.walk():     # only first attachement # !!!
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        print(part.get_content_type())
        filename = part.get_filename()
        file_name, charset = decode_header(filename)[0]
        if charset:
            file_name = file_name.decode(charset)
        print(f"fileName:: {file_name}")
        if bool(file_name):
            with open(folder_to_write / file_name, 'wb') as f:
                f.write(part.get_payload(decode=True))
                # TODO: files_counter  == 1   ??
                break   # only first attachement # !!!
    # TODO: checkfiles_counter ...
    return (folder_to_write, file_name)


def uzip_nd_check(folder_to_write, file_name):
    from zipfile import ZipFile
    import zipfile

    dict_to_return = {
        "exit_is_ok": False,
        "exit_message": "",
        "folder": "",
    }
    #  file_name = ""   # for no zip attached

    if not zipfile.is_zipfile(folder_to_write / file_name):
        dict_to_return["exit_message"] = 'not ZIP file attached'
        return dict_to_return
    # here is ZIP only )
    with ZipFile(folder_to_write / file_name) as zip_file:
        if zip_file.testzip():
            dict_to_return["exit_message"] = f"{file_name} - brocken *.zip"
            return dict_to_return
        else:
            zip_file.extractall(folder_to_write)

    (folder_to_write / file_name).unlink()    # zip_file deleting

    # here we have unpacked files from *.zip
    items_in_folder = [file for file in folder_to_write.iterdir()]
    # if zip has another - zips - unzips them and remove zips
    for item in items_in_folder:
        if zipfile.is_zipfile(item):
            with ZipFile(item) as zip_file:
                zip_file.extractall(folder_to_write)
            item.unlink()    # zip_file deleting

    items_in_folder = [file for file in folder_to_write.iterdir()]
    number_files_in_folder = len(items_in_folder)

    if number_files_in_folder == 1:
        if items_in_folder[0].is_dir():  # in ZIP was folder with files
            # moving files to parent folder
            files_moveup_from(items_in_folder[0])
            items_in_folder = [file for file in folder_to_write.iterdir()]
            number_files_in_folder = len(items_in_folder)
            if number_files_in_folder < 2:
                shutil.rmtree(folder_to_write)
                dict_to_return["exit_message"] = f"Folder {folder_to_write} has less than 2 files from ZIP"
                return dict_to_return

    # FEE maybe 1 *.txt + 1 *.csv
    if (number_files_in_folder == 2) and all(
            [f.suffix.lower() in {'.csv', '.txt'} for f in items_in_folder]):
        dict_to_return["exit_message"] = f"in {folder_to_write} is {number_files_in_folder} files."
        dict_to_return["folder"] = folder_to_write
        dict_to_return["exit_is_ok"] = True
        return dict_to_return

    elif number_files_in_folder == 2:  # Snapshots zip from MacOs maybe
        for el in items_in_folder:
            if el.is_dir():
                if el.name.find("__MACOSX") >= 0:
                    for d in el.iterdir():      # shutil.rmtree(folder_to_write)
                        [f.unlink() for f in d.iterdir()]
                        d.rmdir()
                    el.rmdir()
                else:
                    files_moveup_from(el)
            else:
                dict_to_return["exit_message"] = f"2 items only in {el}."
                return dict_to_return
        # here we are if MacOS and 3 csv files
        mac_files = set(folder_to_write.iterdir())
        mac_files_csv = set(f for f in mac_files if f.suffix.lower() == '.csv')
        mac_files_not_csv = mac_files - mac_files_csv
        if len(mac_files_csv) < 2:
            dict_to_return["exit_message"] = f"in {folder_to_write} is {number_files_in_folder} files."
            dict_to_return["folder"] = folder_to_write
            return dict_to_return

        [f.unlink() for f in mac_files_not_csv]
        dict_to_return["exit_message"] = f"in {folder_to_write} is {number_files_in_folder} files."
        dict_to_return["folder"] = folder_to_write
        dict_to_return["exit_is_ok"] = True
        return dict_to_return

    else:
        for item in folder_to_write.iterdir():
            if item.is_file():
                item.rename(item.parent / ("_" + item.name))
            else:
                dict_to_return["exit_message"] = 'not only files were in ZIP. Folder(s) exists... '
                # TODO: shutil !!!
                return dict_to_return
        else:
            dict_to_return["exit_message"] = f"in {folder_to_write} is {number_files_in_folder} files."
            dict_to_return["folder"] = folder_to_write
            dict_to_return["exit_is_ok"] = True
            return dict_to_return
    assert False, "oooops"


# In[ ]:


def task_type_detect(in_response: dict):
    """Detect task tipe by files name analise.

    1) three *.csv files with adj_ & rec_ & rei_ in names => type "Snapshot"
    2) any quantity files with *.txt extension + only one file with *.csv - extension => type "Fee"

    Args_:
        response (dict): [description]
    """
    from collections import Counter

    my_response = {         # dict_to_return
        "exit_is_ok": False,
        "exit_message": "",
        "folder": "",
    }
    my_response.update({"task_type": ""})

    my_response["folder"] = in_response["folder"]

    suffix_count = Counter(file.suffix for file in in_response["folder"].iterdir())
    if set(suffix_count) <= {".csv", ".txt"}:
        if (suffix_count[".txt"] == 0) and sum(suffix_count.values()) in [3, 4]:  # Snapshots-like
            my_response["exit_is_ok"] = True
            my_response["task_type"] = "Snapshots"
        elif (suffix_count[".csv"] == 1) & (suffix_count[".txt"] > 0):  # Fee-like
            my_response["exit_is_ok"] = True
            my_response["task_type"] = "FEE"
        else:
            file_names = (file.name for file in in_response["folder"].iterdir())
            out_message = f"wrong files set: {file_names}. Can't detect task type..."
            my_response["exit_message"] = out_message   # wrong set of suffixes
    else:  # wrong suffixes - not only .csv | .txt
        my_response["exit_message"] = "files extension not in {'.csv', '.txt'}"

    return my_response


def package_received(folders_path, email_message):
    """Get ayyachments from email and unzip it in created folder. """
    my_response = {         # dict_to_return
        "exit_is_ok": False,
        "exit_message": "",
        "folder": "",
    }
    #  zip_file_name  -> zips_list
    new_folder_with_work, zip_file_name = get_attachments(email_message, folders_path)

    if zip_file_name == 'was_no_attachements':
        my_response["exit_is_ok"] = False
        my_response["exit_message"] = 'was_no_attachements'
        return my_response

    response = resp_unzip = uzip_nd_check(new_folder_with_work, zip_file_name)
    if response["exit_is_ok"]:
        response = task_type_detect(resp_unzip)
        # create subfolder according task type and move received files in it.
        path_to = create_folder(folders_path,
                                response["task_type"] + "/" + response["folder"].name,
                                )
        path_to.rmdir()
        # removing folder to just created subfolder. Works only on same drive !!!
        response["folder"] = new_folder_with_work.rename(path_to)   # i.e.  = path_to

    return response


# In[ ]:
def get_email(uid):
    """Extract emails from byte array."""
    is_ok, data = imap.uid('fetch', uid, '(RFC822)')
    assert is_ok == "OK", f"Cant fetch {uid}"
    # raw mail (data) returning
    email_message = email.message_from_bytes(data[0][1])
    print(f"{uid}:: {get_body(email_message).decode('latin1')}\n")   # 'utf-8'
    return email_message


def imap_in(user, password, email_folder="INBOX"):
    """Login to email folder using imap."""
    global imap
    imap = imaplib.IMAP4_SSL(imap_url)
    imap.login(user, password)
    imap.select(email_folder)


def imap_out():
    """Logout imap connection."""
    imap.close()
    imap.logout()


def packadges_uids(sender='ychudnovs@gmail.com', email_folder="INBOX"):
    imap_in(user, password, email_folder)
    # ------  'ychudnovs@gmail.com'  'vikolo@i.ua'
    print(f'\n\n-----{sender}-------\n\n')
    picked_uids = searched_emails_uids(f'(FROM "{sender}" UNSEEN)', imap)
    return picked_uids


def _send_answer(original, file_to_attach, body_text="Macros reply", copy_to=[]):
    # https://stackoverflow.com/questions/2182196/how-do-i-reply-to-an-email-using-the-python-imaplib-and-include-the-original-mes

    import smtplib
    # First, replace all the attachments
    # in the original message with text/plain placeholders:
    for part in original.walk():
        if (part.get('Content-Disposition')
                and part.get('Content-Disposition').startswith("attachment")):
            part.set_type("text/plain")
            part.set_payload("Attachment removed: %s (%s, %d bytes)"
                             % (part.get_filename(),
                                part.get_content_type(),
                                len(part.get_payload(decode=True))))
            del part["Content-Disposition"]
            del part["Content-Transfer-Encoding"]

    # Then create a reply message:
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    # from email.mime.message import MIMEMessage
    from email.mime.base import MIMEBase

    new = MIMEMultipart("mixed")
    body = MIMEMultipart("alternative")
    body.attach(MIMEText(body_text, "plain"))
    body.attach(MIMEText(f"<html> {body_text} </html>", "html"))
    new.attach(body)

    new["Message-ID"] = email.utils.make_msgid()
    new["In-Reply-To"] = original["Message-ID"]
#    new["References"] = original["Message-ID"]
    new["Subject"] = "Re: " + original["Subject"]
    new["To"] = original["From"].lower()
    new["From"] = user

    all_addrs = new["To"]
    if copy_to:
        copy_to = [addr.lower() for addr in copy_to]
        # якщо адресат є у списку адресів для копії листа, то ми видадяємл його з копії
        copy_to = list(set(copy_to) - set([new["To"], ]))
        if copy_to:
            new['Cc'] = ",".join(copy_to)
            all_addrs += "," + new['Cc']

    #  Then attach the original MIME message object and send:
    # email.encoders.encode_base64(original)
    # new.attach(MIMEMessage(original))

    #  Then attach the zip-file :  file_to_attach
    file_name = file_to_attach.name
    part = MIMEBase('application', 'octet-stream')
    with open(file_to_attach, 'rb') as attachment:
        part.set_payload(attachment.read())
    part.add_header('Content-Disposition', 'attachment', filename=file_name)
    email.encoders.encode_base64(part)
    new.attach(part)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    # s.set_debuglevel(1)
    s.starttls()
    s.login(user, password)
    # s.sendmail(user, [new["To"]], new.as_string())
    s.sendmail(user, all_addrs, new.as_string())
    all_addrs
    s.quit()


def __send_answer(original, file_to_attach, body_text="Macros reply", cc_addrs: list = []):
    # https://stackoverflow.com/questions/2182196/how-do-i-reply-to-an-email-using-the-python-imaplib-and-include-the-original-mes

    import smtplib
    import ssl
    # First, replace all the attachments in the original message with text/plain placeholders:
    for part in original.walk():
        if (part.get('Content-Disposition')
                and part.get('Content-Disposition').startswith("attachment")):
            part.set_type("text/plain")
            part.set_payload("Attachment removed: %s (%s, %d bytes)"
                             % (part.get_filename(),
                                part.get_content_type(),
                                len(part.get_payload(decode=True))))
            del part["Content-Disposition"]
            del part["Content-Transfer-Encoding"]
            print("I delete")

    # Then create a reply message:
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase

    from email.mime.message import MIMEMessage
    from email.utils import formatdate

    new = MIMEMultipart("mixed")
    body = MIMEMultipart("alternative")
    body.attach(MIMEText(body_text, "plain"))
    body.attach(MIMEText(f"<html> {body_text} </html>", "html"))
    new.attach(body)

    new["Message-ID"] = email.utils.make_msgid()
    new["In-Reply-To"] = original["Message-ID"]
    #  new["References"] = original["Message-ID"]
    new["Subject"] = "Re: " + original["Subject"]
    new["To"] = original["From"].lower()
    new["From"] = user

    if cc_addrs:
        new['Cc'] = ",".join(cc_addrs)

    #  Then attach the original MIME message object and send:
    #  new.attach(MIMEMessage(original, ))

    #  Then attach the zip-file :  file_to_attach
    file_name = file_to_attach.name
    part = MIMEBase('application', 'octet-stream')
    with open(file_to_attach, 'rb') as attachment:
        part.set_payload(attachment.read())
    email.encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=file_name)
    new.attach(part)

    context = ssl.create_default_context()
    with smtplib.SMTP('smtp.gmail.com', 587) as s:   # creates SMTP session
        s.ehlo()                            # Can be omitted
        s.starttls(context=context)         # start TLS for security
        s.ehlo()                            # Can be omitted
        s.login(user, password)  # Authentication  # "Password_of_the_sender"
        s.set_debuglevel(True)
        s.send_message(new)


def send_answer(original, file_to_attach: Path, body_text="Macros reply", cc_addrs: list = []):
    # https://stackoverflow.com/questions/2182196/how-do-i-reply-to-an-email-using-the-python-imaplib-and-include-the-original-mes

    import smtplib
    import ssl

    # Then create a reply message:
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase

    new = MIMEMultipart("mixed")
    body = MIMEMultipart("alternative")
    body.attach(MIMEText(body_text, "plain"))
    body.attach(MIMEText(f"<html> {body_text} </html>", "html"))
    new.attach(body)

    new["Message-ID"] = email.utils.make_msgid()
    new["In-Reply-To"] = original["Message-ID"]
    new["References"] = original["Message-ID"]
    new["Subject"] = "Re: " + original["Subject"]
    new["To"] = original["From"].lower()
    new["From"] = user

    if cc_addrs:
        new['Cc'] = ",".join(cc_addrs)

    #  Then attach the file:  file_to_attach
    file_name = file_to_attach.name
    part = MIMEBase('application', 'octet-stream')

    #  with file_to_attach.open(mode='r') as attachment:
    with open(str(file_to_attach), 'rb') as attachment:
        part.set_payload(attachment.read())
    email.encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=file_name)
    new.attach(part)

    context = ssl.create_default_context()
    with smtplib.SMTP('smtp.gmail.com', 587) as s:   # creates SMTP session
        s.ehlo()                            # Can be omitted
        s.starttls(context=context)         # start TLS for security
        s.ehlo()                            # Can be omitted
        s.login(user, password)  # Authentication  # "Password_of_the_sender"
        s.set_debuglevel(False)
        s.send_message(new)


def __send_answer(original, file_to_attach, body_text="Macros reply", cc_addrs: list = []):
    # https://stackoverflow.com/questions/2182196/how-do-i-reply-to-an-email-using-the-python-imaplib-and-include-the-original-mes

    import smtplib
    import ssl

    # First, replace all the attachments in the original message with text/plain placeholders:
    for part in original.walk():
        if (part.get('Content-Disposition')
                and part.get('Content-Disposition').startswith("attachment")):
            part.set_type("text/plain")
            part.set_payload("Attachment removed: %s (%s, %d bytes)"
                             % (part.get_filename(),
                                part.get_content_type(),
                                len(part.get_payload(decode=True))))
            del part["Content-Disposition"]
            del part["Content-Transfer-Encoding"]
            print("I delete")

    # Then create a reply message:
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase

    from email.mime.message import MIMEMessage

    new = MIMEMultipart("mixed")
    body = MIMEMultipart("alternative")
    body.attach(MIMEText(body_text, "plain"))
    body.attach(MIMEText(f"<html> {body_text} </html>", "html"))
    new.attach(body)

    new["Message-ID"] = email.utils.make_msgid()
    new["In-Reply-To"] = original["Message-ID"]
    new["References"] = original["Message-ID"]
    new["Subject"] = "Re: " + original["Subject"]
    new["To"] = original["From"].lower()
    new["From"] = user

    if cc_addrs:
        new['Cc'] = ",".join(cc_addrs)

    #  Then attach the original MIME message object and send:
    new.attach(MIMEMessage(original, ))

    #  Then attach the file:  file_to_attach
    file_name = file_to_attach.name
    part = MIMEBase('application', 'octet-stream')
    with open(file_to_attach, 'rb') as attachment:
        part.set_payload(attachment.read())
    email.encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=file_name)
    new.attach(part)

    context = ssl.create_default_context()
    with smtplib.SMTP('smtp.gmail.com', 587) as s:   # creates SMTP session
        s.ehlo()                            # Can be omitted
        s.starttls(context=context)         # start TLS for security
        s.ehlo()                            # Can be omitted
        s.login(user, password)  # Authentication  # "Password_of_the_sender"
        s.set_debuglevel(False)
        s.send_message(new)


def send_mail_with_attach(
        to_addr: str, mail_subject: str,
        cc_addrs: list = [], bcc_addrs: list = [],
        mail_body: str = "by Python script",
        files_to_attach: list = [Path(r"D:\OneDrive\PyCodes\SHEDULER\tenor_Mister_Bin.gif"), ]
        # TODO:  del abs path
):
    # Python code to Send mail with attachments
    # from your Gmail account
    # https://stackoverflow.com/questions/26582811/gmail-python-multiple-attachments
    # https://realpython.com/python-send-email/
    # https://code.tutsplus.com/ru/tutorials/sending-emails-in-python-with-smtp--cms-29975
    #
    # https://stackoverflow.com/questions/1546367/python-how-to-send-mail-with-to-cc-and-bcc
    #
    # error codes:
    # https://docs.microsoft.com/ru-ru/exchange/mail-flow/test-smtp-with-telnet?view=exchserver-2019

    # libraries to be imported
    import smtplib
    import ssl
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    from email.utils import formatdate

    msg = MIMEMultipart()           # instance of MIMEMultipart

    msg['From'] = user              # storing the senders email address
    msg['To'] = to_addr             # storing the receivers email address
    msg['Subject'] = mail_subject   # storing the subject # "Subject of the Mail"
    msg["Date"] = formatdate(localtime=True)
    msg.preamble = 'Ukraine'

    if cc_addrs:
        msg['Cc'] = ",".join(cc_addrs)
    if bcc_addrs:
        msg['Bcc'] = ",".join(bcc_addrs)

    msg.attach(MIMEText(mail_body, 'plain'))  # "Body_of_the_mail"

    try:
        for file in files_to_attach:
            part = MIMEBase('application', 'octet-stream')
            with open(file, "rb") as fh:  # open the file to be sent
                data = fh.read()
            filename = file.name
            part.set_payload(data)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
            msg.attach(part)
    except IOError:
        msg = f"Error opening attachment file {file.name}"
        print(msg)
        sys.exit(1)

    context = ssl.create_default_context()
    with smtplib.SMTP('smtp.gmail.com', 587) as s:   # creates SMTP session
        s.ehlo()                            # Can be omitted
        s.starttls(context=context)         # start TLS for security
        s.ehlo()                            # Can be omitted
        s.login(user, password)  # Authentication  # "Password_of_the_sender"
        s.set_debuglevel(False)
        s.send_message(msg)
