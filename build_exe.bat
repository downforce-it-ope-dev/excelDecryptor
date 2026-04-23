@echo off
set PYTHON_EXE=C:\Users\downforceITkkt\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe

"%PYTHON_EXE%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  --name ExcelDecryptor ^
  --add-data "templates;templates" ^
  --add-data "update_config.json;." ^
  launcher.py

echo.
echo Build finished.
echo EXE path: dist\ExcelDecryptor.exe
