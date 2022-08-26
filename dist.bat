@echo OFF
echo Begin build...

set PYTHONOPTIMIZE=2 && pyinstaller -D -n "Spreadsheet-Merger" -w -y "./dist.py"

echo Build complete.