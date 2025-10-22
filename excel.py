from io import BytesIO, StringIO
from openpyxl import Workbook
import openpyxl
from openpyxl.descriptors import String

import pandas as pd
from tabulate import tabulate

from convertor import convertToJTSK

START_ROW = 3
RECORD_INDEX_COL = 7
LATITUDE_COL = 9
LONGITUDE_COL = 10
HEIGHT_COL = 34
FEATURE_DESCRIPTION_COL = 47

class Record:
    index: String
    latitude: String
    longitude: String
    description: String
    h: String

def process_sheet(file_location: String, save_directory: String, negative_xy: bool, convert_height: bool) -> bool: # returns success
    if not isinstance(file_location, str) or not isinstance(save_directory, str):
        return False

    wb = openpyxl.load_workbook(file_location)
    sheet = wb.active
    if sheet is None:
        return False

    i = START_ROW
    val = sheet.cell(row=i, column=1).value
    while val is not None:
        i+=1
        val = sheet.cell(row=i, column=1).value

    END_ROW = i-1

    records = []
    longitude_max_len = 0
    latitude_max_len = 0
    index_max_len = 0

    # process input sheet
    for j in range(START_ROW, END_ROW + 1):
        record = Record()

        record.index = sheet.cell(row=j, column=RECORD_INDEX_COL).value

        if len(str(record.index)) > index_max_len:
            index_max_len = len(str(record.index))

        # get lat/longtit and translate
        latitude = float(str(sheet.cell(row=j, column=LATITUDE_COL).value).replace(",", "."))
        longitude = float(str(sheet.cell(row=j, column=LONGITUDE_COL).value).replace(",", "."))

        height = float(str(sheet.cell(row=j, column=HEIGHT_COL).value).replace(",", "."))

        if len(str(latitude)) > latitude_max_len:
            latitude_max_len = len(str(latitude))
        if len(str(longitude)) > longitude_max_len:
            longitude_max_len = len(str(longitude))

        if (latitude == 0 or longitude == 0):
            record.latitude = 0
            record.longitude = 0
            record.h = 0
        else:
            e = convertToJTSK(latitude, longitude, height)

            if e is None:
                return False

            record.latitude = abs(e[0]) if not negative_xy else -abs(e[0])
            record.longitude = abs(e[1]) if not negative_xy else -abs(e[1])
            if convert_height:
                record.h = round(e[2], 2) if e[2] is not None else str(0)
            else:
                record.h = height if height is not None else str(0)

        record.description = sheet.cell(row=j, column=FEATURE_DESCRIPTION_COL).value

        records.append(record)

    # sort the records just in case
    sorted(records, key=lambda record: record.index)
    df = pd.DataFrame([(r.index, r.latitude if r.latitude!=0 else str(0), r.longitude if r.longitude!=0 else str(0), r.h if r.h!=0 else str(0), r.description) for r in records]) #, columns=["# Číslo bodu", "Y", "X", "Z", "Popis bodu"])

    file = file_location.split("/")
    file_name = file[-1].split(".")[0]

    with open(save_directory + "/" + file_name + ".txt", "w", encoding="utf-8") as f:

        f.write(df.to_string(index=False, header=False))

    return True
