import pywintypes
from rich.console import Console
import pandas as pd

from connection import outlook_connection_inbox

console = Console(
    log_time_format="[%Y-%m-%d %H:%M:%S]"
)

try:
    folder, outlook = outlook_connection_inbox()

    console.log("Połączono z pocztą Outlook")
except pywintypes.com_error as e:
    console.log(f"Wystąpił błąd podczas łączenia z pocztą Outlook: {e}")
    raise

assert folder is not None
messages = folder.Items
messages.Sort("[ReceivedTime]", True)

for i in messages:
    if i.Class != 43:
        continue
    print(f"{i.HTMLbody}")
    console.rule("[bold red]BREAK")
    console.input(">")

# df =
