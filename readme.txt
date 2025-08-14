pip install pandas openpyxl docxtpl


pyinstaller --clean --onefile --noconsole --icon=logo.ico --add-data "logo.ico;." batch_generate_contracts.py


