::@echo off
::setlocal
::set PYTHON_EXE=C:\Users\micas\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe
::if "%*"=="" (
::    "%PYTHON_EXE%" "%~dp0run_agent.py" all
::) else (
::    "%PYTHON_EXE%" "%~dp0run_agent.py" %*
::)
::pause


@echo off
setlocal
set PYTHON_EXE=python
if "%*"=="" (
    "%PYTHON_EXE%" "%~dp0run_agent.py" all
) else (
    "%PYTHON_EXE%" "%~dp0run_agent.py" %*
)
pause