@echo off
cd /d %~dp0

:: Python �X�N���v�g�̎��s
python fam8_Mixx�i��Report�擾�Ə�������.py
if %errorlevel% neq 0 exit /b

:: Python �X�N���v�g�̎��s
python fam8_AX-AD�i��Report�擾�Ə�������.py
if %errorlevel% neq 0 exit /b


exit