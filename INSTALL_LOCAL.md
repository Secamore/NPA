# Local installation for NPA

## Install on a local device

1. On the host machine, open the app in a browser and keep it running on the local network.
2. In Chrome, Edge, or Safari, use the browser menu and choose Install app or Add to Home Screen.
3. On Windows or macOS, install the app as a desktop shortcut. On Android or iPhone, add it to the home screen.
4. Launch the app from the installed shortcut and connect to the local host address.

## Offline use

When the device has no internet access, the app still works if:

- the host machine is running the app on the private LAN,
- the device connects to the host over Wi-Fi or LAN,
- the host is started with the offline environment settings, and
- the database and assets are stored in a local folder.

## Host setup

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:NPA_OFFLINE = "1"
$env:NPA_MANAGER_PASSWORD = "replace-with-a-strong-password"
$env:NPA_MANAGER_INVITE = "replace-with-a-private-invite-code"
$env:NPA_DB_PATH = "$HOME\NPAData\npa_portal.db"
$env:NPA_ASSETS_DIR = "$HOME\NPAData\assets"
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

macOS/Linux:

```sh
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export NPA_OFFLINE=1
export NPA_MANAGER_PASSWORD='replace-with-a-strong-password'
export NPA_MANAGER_INVITE='replace-with-a-private-invite-code'
export NPA_DB_PATH="$HOME/NPAData/npa_portal.db"
export NPA_ASSETS_DIR="$HOME/NPAData/assets"
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

## Connect devices

Find the host's private LAN IP address and open:

```text
http://<host-ip>:8501
```

Use the browser install flow on each device to keep the app available without internet access.
