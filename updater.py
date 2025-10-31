import os
import urllib.request

SERVER_VERSION_URL = "https://esites.pro/coderdoc/version.txt"
SERVER_SCRIPT_URL = "https://esites.pro/coderdoc/compilar_exames.py"
LOCAL_VERSION_FILE = "version.txt"
LOCAL_SCRIPT_FILE = "compilar_exames.py"

def get_local_version():
    if not os.path.exists(LOCAL_VERSION_FILE):
        return "0.0.0"
    with open(LOCAL_VERSION_FILE, "r") as f:
        return f.read().strip()

def get_server_version():
    try:
        with urllib.request.urlopen(SERVER_VERSION_URL) as response:
            return response.read().decode("utf-8").strip()
    except Exception:
        return None

def update_if_needed():
    local_version = get_local_version()
    server_version = get_server_version()

    if not server_version:
        print("⚠️  No internet or server unavailable. Running offline mode.")
        return False

    if server_version != local_version:
        print(f"⬆️ Updating from {local_version} → {server_version}")
        urllib.request.urlretrieve(SERVER_SCRIPT_URL, LOCAL_SCRIPT_FILE)
        with open(LOCAL_VERSION_FILE, "w") as f:
            f.write(server_version)
        print("✅ Update complete.")
        return True

    print("✔️ Already up to date.")
    return False
