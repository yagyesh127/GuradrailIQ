#!/usr/bin/env bash
# GuardrailIQ - setup + launch (macOS / Linux)
set -e
cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  echo "[setup] creating virtual environment .venv ..."
  python3 -m venv .venv
fi
# shellcheck disable=SC1091
source .venv/bin/activate

echo "[setup] installing dependencies ..."
python -m pip install --upgrade pip >/dev/null
python -m pip install -r requirements.txt

echo "[run] launching GuardrailIQ at http://localhost:8501 (Ctrl+C to stop) ..."
streamlit run app/guardrailiq_app.py
