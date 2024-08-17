# EXE to delete Specified field in Excel file

* Install the Required Libraries
```
pip install pandas openpyxl tk pyinstaller
```


* Use PyInstaller to Create the Executable
```
pyinstaller --onefile --windowed .\DeletionOfID.py --distpath . --name [your EXE file name]
```