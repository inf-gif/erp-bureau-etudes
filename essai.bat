@echo off
title Gestion des Offres - Serveur
color 0A
cd /d "%~dp0"
call venv\Scripts\activate.bat
echo Demarrage du serveur...
python app_web.py
pause