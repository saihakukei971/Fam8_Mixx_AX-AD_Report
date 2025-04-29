@echo off
cd /d %~dp0

:: Python スクリプトの実行
python fam8_Mixx進捗Report取得と書き込み.py
if %errorlevel% neq 0 exit /b

:: Python スクリプトの実行
python fam8_AX-AD進捗Report取得と書き込み.py
if %errorlevel% neq 0 exit /b


exit