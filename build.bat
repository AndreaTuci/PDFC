pyinstaller --noconfirm --log-level=WARN ^
    --onefile --windowed ^
    --add-data=".\img\header_s.png;img" ^
    --icon=.\img\app.ico ^
    PDFC.py