@echo off
cd /d %~dp0

call "C:\ProgramData\anaconda3\Scripts\activate.bat" voxcpm

start cmd /c "timeout /t 5 >nul & start http://localhost:8808"

python app.py --port 8808

pause