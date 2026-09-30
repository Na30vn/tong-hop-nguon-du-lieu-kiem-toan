@echo off
setlocal
cd /d "%~dp0"
"C:\Users\anhvn1\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" build_data.py
if errorlevel 1 (
  echo Cap nhat khong thanh cong.
  pause
  exit /b 1
)
echo Da cap nhat du lieu dashboard trong thu muc dist.
pause
