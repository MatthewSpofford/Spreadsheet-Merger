@echo OFF
echo Begin build...

python -OO -m pyinstaller -F -n "Spreadsheet-Merger" -w -y "./dist.py"

echo Build complete.