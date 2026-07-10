::@echo off
::setlocal
::set PYTHON_EXE=C:\Users\micas\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe
::"%PYTHON_EXE%" "%~dp0run_weekly_report.py"
::#pause

@echo off
cd /d "%~dp0"
python run_weekly_report.py
pause
