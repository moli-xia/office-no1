@echo off
cd /d %~dp0
echo == 材料排版工具 · 打包为单个 exe ==
py -m PyInstaller --noconfirm --clean --onefile --windowed --collect-all tkinterdnd2 --collect-all docx --name "材料排版工具" app.py
if errorlevel 1 (
  echo 打包失败，请确认已安装: py -m pip install python-docx pyinstaller tkinterdnd2
  pause
  exit /b 1
)
echo.
echo [OK] 打包完成: dist\材料排版工具.exe
pause
