@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
if not exist ".venv\Scripts\python.exe" (
  echo Preparando o ambiente pela primeira vez...
  py -3 -m venv .venv 2>nul || python -m venv .venv
  ".venv\Scripts\python.exe" -m pip install --quiet openpyxl
)
title Painel de contas a pagar - deixe esta janela aberta
".venv\Scripts\python.exe" painel.py
pause
