@echo OFF
echo Begin build...

pyinstaller -F -n "Spreadsheet Merger" -w --python-option "-O" -y "./dist.py"

echo Build complete.