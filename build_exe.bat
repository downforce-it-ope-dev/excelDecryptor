@echo off
set PYTHON_EXE=C:\Users\downforceITkkt\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe

if exist updater_dist rmdir /s /q updater_dist
if exist updater_build rmdir /s /q updater_build

"%PYTHON_EXE%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  --name dftsExcelDecryptorUpdater ^
  --distpath updater_dist ^
  --workpath updater_build ^
  updater_app.py

"%PYTHON_EXE%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  --name dftsExcelDecryptor ^
  --add-binary "updater_dist\dftsExcelDecryptorUpdater.exe;." ^
  --add-data "templates;templates" ^
  --add-data "update_config.json;." ^
  launcher.py

echo.
echo Build finished.
echo EXE path: dist\dftsExcelDecryptor.exe
