import os

import pyodbc
import win32com.client
from dotenv import load_dotenv

load_dotenv()

SERVER = os.getenv('SERVER')
DATABASE = os.getenv('DATABASE')
UID = os.getenv('UID')
PWD = os.getenv('PWD')


def sql_connection_open():
    conn = pyodbc.connect(
        "DRIVER={ODBC Driver 18 for SQL Server};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        f"UID={UID};"
        f"PWD={PWD};"
        "TrustServerCertificate=yes;"
    )
    return conn


def outlook_connection():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    recipient = namespace.CreateRecipient("Pomoc IT")
    recipient.Resolve()
    mailbox = namespace.Folders["Pomoc IT"]
    inbox = mailbox.Folders["Skrzynka odbiorcza"]
    folder = inbox.Folders["Reset Microsoft"]
    return folder, outlook

def outlook_connection_pomoc():

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)
    folder = inbox.Folders["Reset Microsoft"]

    return folder, outlook
