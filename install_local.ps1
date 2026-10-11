$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

python -m venv .venv
. .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

$env:NPA_OFFLINE = "1"
$env:NPA_MANAGER_PASSWORD = if ($env:NPA_MANAGER_PASSWORD) { $env:NPA_MANAGER_PASSWORD } else { "ChangeMe123!" }
$env:NPA_MANAGER_INVITE = if ($env:NPA_MANAGER_INVITE) { $env:NPA_MANAGER_INVITE } else { "NPA-MGR-2026" }
$env:NPA_DB_PATH = if ($env:NPA_DB_PATH) { $env:NPA_DB_PATH } else { "$HOME\NPAData\npa_portal.db" }
$env:NPA_ASSETS_DIR = if ($env:NPA_ASSETS_DIR) { $env:NPA_ASSETS_DIR } else { "$HOME\NPAData\assets" }

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $env:NPA_DB_PATH) | Out-Null
New-Item -ItemType Directory -Force -Path $env:NPA_ASSETS_DIR | Out-Null
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
