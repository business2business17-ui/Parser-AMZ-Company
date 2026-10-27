@echo off
setlocal
pushd "%~dp0"
where py >nul 2>nul
if errorlevel 1 goto use_python
py -3 setup_env.py --run %*
goto finished
:use_python
python setup_env.py --run %*
:finished
set "XRAY_EXIT=%ERRORLEVEL%"
popd
echo.
if not "%XRAY_EXIT%"=="0" echo Run failed or collected no sellers. Read the message above and reports.
pause
exit /b %XRAY_EXIT%
