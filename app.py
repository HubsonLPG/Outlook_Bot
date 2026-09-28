import re
import time

import pyodbc
import pywintypes

from connection import outlook_connection, sql_connection_open

try:
    folder, outlook = outlook_connection()
    print("Połączono z pocztą Outlook")
except pywintypes.com_error as e:
    print(f"Wystąpił błąd podczas łączenia z pocztą Outlook: {e}")
    raise

try:
    conn = sql_connection_open()
    cursor = conn.cursor()
    print("Połączono z bazą SQL")
except pyodbc.Error as e:
    print(f"Wystąpił problem z łączeniem z bazą danych: {e}")
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

                print(f"Przetwarzam: {email}")
                send_email(row.OS_EMAIL2, mail_body, True)
                message.FlagStatus = 1
                message.UnRead = False
                message.Save()

                print("Wiadomość obsłużona!")

            else:
                print(f"Nie znaleziono adresu dla: {email}")

        except Exception as e:  # noqa: BLE001
            print(f"Błąd przy obsłudze {email}: {e}")


while True:

    try:
        app(folder, conn)

    except Exception as e:  # noqa: BLE001
        print(f"Błąd sprawdzania skrzynki: {e}")

    time.sleep(30)
