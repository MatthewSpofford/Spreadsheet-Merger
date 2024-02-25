@echo OFF
echo Begin build...

call .\venv\Scripts\activate.bat
set PYTHONOPTIMIZE=2 && pyinstaller -D -n "Spreadsheet-Merger" -w -y "./dist_entrypoint.pyw"
call deactivate

echo Build complete.