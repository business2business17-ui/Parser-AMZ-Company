@echo off
setlocal
title Amazon Parser - USA 90001
pushd "%~dp0"
set "XRAY_RESUME=%~1"
if defined XRAY_RESUME goto resume
for /f "delims=" %%D in ('dir /b /ad /o-n "reports\20*_UTC" 2^>nul') do if not defined XRAY_RESUME set "XRAY_RESUME=reports\%%D"
if defined XRAY_RESUME goto resume
call run.bat --no-headless --zip US=90001 --skip-storefront
goto finished
:resume
echo Resuming: "%XRAY_RESUME%"
echo Delivery: USA, Los Angeles 90001
echo Complete Amazon verification in the browser. No automatic timeout.
call run.bat --no-headless --zip US=90001 --skip-storefront --resume "%XRAY_RESUME%"
:finished
set "XRAY_EXIT=%ERRORLEVEL%"
popd
exit /b %XRAY_EXIT%
