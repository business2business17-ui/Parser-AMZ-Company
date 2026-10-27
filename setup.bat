@echo off
setlocal
pushd "%~dp0"
where py >nul 2>nul
if errorlevel 1 goto use_python
py -3 setup_env.py
goto finished
:use_python
python setup_env.py
:finished
set "XRAY_EXIT=%ERRORLEVEL%"
popd
pause
exit /b %XRAY_EXIT%
