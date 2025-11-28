#!/usr/bin/env bash
# run_mac.sh
# Simple runner: activates the venv created by mac/installer.sh and runs the app as a module
# Usage (from project root): chmod +x run_mac.sh && ./run_mac.sh

set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "$0")" && pwd)"
VENV_DIR="$PROJECT_ROOT/.venv_mac"

# Check venv
if [ ! -d "$VENV_DIR" ]; then
  echo "Virtualenv não encontrado em $VENV_DIR."
  echo "Execute primeiro: ./mac/installer.sh"
  exit 1
fi

# Activate venv
# shellcheck source=/dev/null
source "$VENV_DIR/bin/activate"

# Run application as module so package imports work
python -m src.main

# deactivate on exit
deactivate || true