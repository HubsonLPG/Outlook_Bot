import pyodbc
import win32com.client


def sql_connection_open():
    conn = pyodbc.connect(
        "DRIVER={ODBC Driver 18 for SQL Server};"
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
