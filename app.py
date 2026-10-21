import re
import time

import pyodbc
import pywintypes
from rich.console import Console

from connection import outlook_connection, outlook_connection_pomoc, sql_connection_open  # noqa: F401

console = Console(
    log_time_format="[%Y-%m-%d %H:%M:%S]"
)

try:
    folder, outlook = outlook_connection_pomoc()
    # folder, outlook = outlook_connection()
    console.log("Połączono z pocztą Outlook")
except pywintypes.com_error as e:
    console.log(f"Wystąpił błąd podczas łączenia z pocztą Outlook: {e}")
    raise

try:
    conn = sql_connection_open()
    cursor = conn.cursor()
    console.log("Połączono z bazą SQL")
except pyodbc.Error as e:
    console.log(f"Wystąpił problem z łączeniem z bazą danych: {e}")
    raise

mail_body = (
    """
<html>
<body>
    <p>Otrzymaliśmy prośbę o reset hasła z konta użytkownika.</p>

    <p>
        Jeśli taka potrzeba została zgłoszona przez panel
        resetu hasła na platformie Office, prosimy skorzystać ze strony -
        <a href="https://wu.wsb.edu.pl/wu/odzyskiwaniehasla.aspx">
            https://wu.wsb.edu.pl/wu/odzyskiwaniehasla.aspx
        </a>
    </p>
</body>
</html>
"""
)


def send_email(email_address, mail_body, send=False):
    assert outlook is not None
    mail = outlook.CreateItem(0)

    mail.To = str(email_address)
    mail.Subject = "Prośba o reset hasła"
    mail.BodyFormat = 2
    mail.HTMLBody = mail_body
    mail.SentOnBehalfOfName = "pomoc.office365@wsb.edu.pl"

    if send == True:
        mail.Send()
    elif send == False:
        mail.Display()


def app(folder, conn):

    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)

    for message in messages:

        if message.FlagStatus == 1:
            continue

        email = re.search(
            r'[\w\.-]+@[\w\.-]+\.\w+',
            message.Subject
        ).group()  # type: ignore

        try:
            # cursor = conn.cursor()

            cursor.execute("""
                SELECT OS_EMAIL, OS_EMAIL2
                FROM Osoba
                WHERE OS_EMAIL = ?
            """, email)

            row = cursor.fetchone()

            if row and row.OS_EMAIL2:

                console.log(f"Przetwarzam: {email}")
                send_email(row.OS_EMAIL2, mail_body, True)
                message.FlagStatus = 1
                message.UnRead = False
                message.Save()

                console.log("Wiadomość obsłużona!")

            else:
                console.log(f"Nie znaleziono adresu dla: {email}")

        except Exception as e:  # noqa: BLE001
            console.log(f"Błąd przy obsłudze {email}: {e}")


while True:
    try:
        app(folder, conn)

    except pywintypes.com_error as e:
        console.log(f"[red]Błąd Outlooka:[/red] {e}")
        console.log("Próbuję ponownie połączyć się z Outlookiem...")

        try:
            folder, outlook = outlook_connection_pomoc()
            console.log("[green]Ponownie połączono z Outlookiem[/green]")

        except pywintypes.com_error as reconnect_error:
            console.log(
                f"[red]Nie udało się ponownie połączyć:[/red] "
                f"{reconnect_error}"
            )

    except Exception as e:  # noqa: BLE001
        console.log(f"[red]Błąd sprawdzania skrzynki:[/red] {e}")

    time.sleep(30)
