@echo off
rem Double-click to start the web page:   kvc_web   ==   python kvc_web.py
python "%~dp0kvc_web.py" %*
if errorlevel 1 pause
