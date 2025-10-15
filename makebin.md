py -m pip install -r requirements.txt
py -m PyInstaller --onefile --add-data "./table_yx_3_v1710.dat:." --add-data "./CR-2005_v1005.dat:." --noconsole main.py
