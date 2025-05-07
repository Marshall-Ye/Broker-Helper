from pyupdater.client import Client
from client_config import ClientConfig
import threading

__version__ = "1.3.4"  # Make sure this matches your built version

def check_for_updates_ui(callback=None):
    def worker():
        client = Client(ClientConfig(), refresh=True)
        client.refresh()
        update = client.update_check("GABrokerToolkit", __version__)
        if update:
            update.download()
            if update.is_downloaded():
                update.extract_restart()
                if callback:
                    callback("Update downloaded. Restarting app...")
        else:
            if callback:
                callback("You're already on the latest version.")

    threading.Thread(target=worker, daemon=True).start()
