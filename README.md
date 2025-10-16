# WGS84 to S-JTSK

This tool has been made to convert xlsx files of specific format with WGS84 coordinates to a specific txt format with S-JTSK coordinates.

Data file sources:
[CR-2005_v1005](https://cuzk.gov.cz/Zememerictvi/Geodeticke-zaklady-na-uzemi-CR/GNSS/Nova-realizace-systemu-ETRS89-v-CR/CR-2005_v1005.aspx)

[table_yx_3_v1710](https://cuzk.gov.cz/Zememerictvi/Geodeticke-zaklady-na-uzemi-CR/GNSS/Nova-realizace-systemu-ETRS89-v-CR/table_yx_3_v1710.aspx)

---

## Usage
Install the required libraries:
```
pip install -r requirements.txt
```
Run the app:
```
python main.py
```

## Building self-contained executable
I recommend creating a virtual environment and using it to build the executable, this will make the final package smaller:
```
pip install pyinstaller
python3 -m PyInstaller --onefile --add-data "./table_yx_3_v1710.dat:." --add-data "./CR-2005_v1005.dat:." --noconsole main.py
```
