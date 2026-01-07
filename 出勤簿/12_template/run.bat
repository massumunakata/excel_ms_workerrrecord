@echo off
setlocal

REM バッチファイルのあるフォルダを基準にする
cd /d "%~dp0"

REM PowerShellスクリプトを実行
powershell -ExecutionPolicy Bypass -File "%~dp0MS_salary.ps1"

endlocal
