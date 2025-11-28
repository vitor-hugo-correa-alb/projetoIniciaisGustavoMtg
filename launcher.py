#!/usr/bin/env python3
# launcher.py — executa o módulo src.main como script (preserva imports relativos)
# IMPORT ESTÁTICO para o PyInstaller detectar o package `src`
import src.main  # for PyInstaller static analysis / inclusion
import runpy

if __name__ == "__main__":
    # Executa src.main como __main__ (equivalente a python -m src.main)
    runpy.run_module("src.main", run_name="__main__")