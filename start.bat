@echo off
cd /d "%~dp0"
python pandoc_gui_converter.py
if errorlevel 1 pause
