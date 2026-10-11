#!/usr/bin/env bash
set -eu

cd "$(dirname "$0")"

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

export NPA_OFFLINE=1
export NPA_MANAGER_PASSWORD="${NPA_MANAGER_PASSWORD:-ChangeMe123!}"
export NPA_MANAGER_INVITE="${NPA_MANAGER_INVITE:-NPA-MGR-2026}"
export NPA_DB_PATH="${NPA_DB_PATH:-$HOME/NPAData/npa_portal.db}"
export NPA_ASSETS_DIR="${NPA_ASSETS_DIR:-$HOME/NPAData/assets}"

mkdir -p "$(dirname "$NPA_DB_PATH")" "$NPA_ASSETS_DIR"
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
