pyinstaller --onefile --windowed --icon=czf.ico czf.py
xcopy /-i /y dist\czf.exe czf.exe
rmdir /s /q dist
rmdir /s /q build
del czf.spec