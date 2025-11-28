#!/usr/bin/env bash
# mac/installer.sh
# Create a mac-friendly virtualenv and install project dependencies.
# Usage (from project root): chmod +x mac/installer.sh && ./mac/installer.sh

set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
VENV_DIR="$PROJECT_ROOT/.venv_mac"
REQ_FILE="$PROJECT_ROOT/mac/requirements.txt"

echo "Projeto: $PROJECT_ROOT"
echo "Virtualenv: $VENV_DIR"
echo "Requirements: $REQ_FILE"
echo

# Ensure python3 is available
if ! command -v python3 >/dev/null 2>&1; then
  echo "Erro: python3 não encontrado. Instale o Python 3 (recomendo 3.10+)."
  echo "No mac: https://www.python.org/downloads/ ou 'brew install python'"
  exit 1
fi

PYTHON_BIN="$(command -v python3)"
echo "Usando Python em: $PYTHON_BIN"
echo

# Create venv if missing
if [ ! -d "$VENV_DIR" ]; then
  echo "Criando virtualenv em $VENV_DIR ..."
  "$PYTHON_BIN" -m venv "$VENV_DIR"
fi

# Activate venv
# shellcheck source=/dev/null
source "$VENV_DIR/bin/activate"

# Upgrade packaging tools
echo "Atualizando pip, setuptools e wheel..."
python -m pip install --upgrade pip setuptools wheel

# Install requirements
if [ -f "$REQ_FILE" ]; then
  echo "Instalando dependências via pip (isso pode demorar)..."
  if ! pip install -r "$REQ_FILE"; then
    echo
    echo "Instalação falhou — provavelmente dependências nativas como lxml/Pillow precisam de libs do sistema."
    echo "Siga estas instruções e rode este script novamente:"
    echo "  1) Instale Xcode Command Line Tools: xcode-select --install"
    echo "  2) Instale dependências via Homebrew (recomendo Homebrew para macOS):"
    echo "       brew install libxml2 libxslt pkg-config zlib libjpeg"
    echo "     (on Apple Silicon brew path may be /opt/homebrew)"
    echo "  3) Export flags se necessário, ex.:"
    echo "       export LDFLAGS=\"-L/opt/homebrew/opt/libxml2/lib -L/opt/homebrew/opt/libxslt/lib\""
    echo "       export CPPFLAGS=\"-I/opt/homebrew/opt/libxml2/include -I/opt/homebrew/opt/libxslt/include\""
    echo "  4) Reinstale requisitos: pip install -r $REQ_FILE"
    deactivate || true
    exit 1
  fi
else
  echo "Arquivo de requirements não encontrado: $REQ_FILE"
  deactivate || true
  exit 1
fi

echo
echo "Dependências instaladas com sucesso no venv: $VENV_DIR"
echo "Para executar a aplicação: ./run_mac.sh (na raiz do projeto)."
echo "Se preferir, ative o venv manualmente:"
echo "  source \"$VENV_DIR/bin/activate\""
echo "  python -m src.main"
deactivate || true