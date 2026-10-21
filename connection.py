import os

import pyodbc
import win32com.client
from dotenv import load_dotenv
from rich.console import Console

console = Console(
    log_time_format="[%Y-%m-%d %H:%M:%S]"
)

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
    step = "Uruchamianie Outlooka"

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")

        step = "Pobieranie przestrzeni MAPI"
        namespace = outlook.GetNamespace("MAPI")

        step = "Otwieranie skrzynki Pomoc IT"
        mailbox = namespace.Folders["Pomoc IT"]

        step = "Otwieranie skrzynki odbiorczej"
        inbox = mailbox.Folders["Skrzynka odbiorcza"]

        step = "Otwieranie folderu Reset Microsoft"
        folder = inbox.Folders["Reset Microsoft"]

        return folder, outlook

    except Exception as e:  # noqa: BLE001
        console.log(f"[red]Błąd ({step}): {e}[/red]")
        return None, None


def outlook_connection_pomoc():
    step = "Uruchamianie Outlooka"

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")

        step = "Pobieranie przestrzeni MAPI"
        namespace = outlook.GetNamespace("MAPI")

        step = "Otwieranie skrzynki odbiorczej"
        inbox = namespace.GetDefaultFolder(6)

        step = "Otwieranie folderu Reset Microsoft"
        folder = inbox.Folders["Reset Microsoft"]

        return folder, outlook

    except Exception as e:  # noqa: BLE001
        console.log(f"[red]Błąd ({step}): {e}[/red]")
        return None, None


def outlook_connection_inbox():
    step = "Uruchamianie Outlooka"

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")

        step = "Pobieranie przestrzeni MAPI"
        namespace = outlook.GetNamespace("MAPI")

        step = "Otwieranie skrzynki Pomoc IT"
        mailbox = namespace.Folders["Pomoc IT"]

        step = "Otwieranie skrzynki odbiorczej"
        folder = mailbox.Folders["Skrzynka odbiorcza"]

        return folder, outlook

    except Exception as e:  # noqa: BLE001
        console.log(f"[red]Błąd ({step}): {e}[/red]")
        return None, None
