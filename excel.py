from io import BytesIO, StringIO
from openpyxl import Workbook
import openpyxl
from openpyxl.descriptors import String

import pandas as pd
from tabulate import tabulate

from convertor import convertToJTSK

START_ROW = 3
HEADER_ROW = 1

DEFAULT_RECORD_INDEX_COLUMN = "Record Index"
DEFAULT_LATITUDE_COLUMN = "Latitude"
DEFAULT_LONGITUDE_COLUMN = "Longitude"
DEFAULT_DEPTH_COLUMN = "Depth (m)"
DEFAULT_HEIGHT_COLUMN = "Altitude/Locator Antenna/Geoid (m)"
DEFAULT_DESCRIPTION_COLUMN = "Feature Description"

def find_column(sheet, name: String) -> int | None:
    for col in range(1, sheet.max_column + 1):
        if str(sheet.cell(row=HEADER_ROW, column=col).value).strip() == str(name).strip():
            return col
    return None

def missing_column(file_location: String, columns) -> String | None: # returns the first name not found in the header, or None
    wb = openpyxl.load_workbook(file_location)
    sheet = wb.active
    if sheet is None:
        return columns[0]
    for name in columns:
        if find_column(sheet, name) is None:
            return name
    return None

class Record:
    index: String
    latitude: String
    longitude: String
    description: String
    height: String
    depth: String

def process_sheet(file_location: String, save_directory: String, negative_xy: bool, convert_height: bool, subtract_depth: bool,
                  record_index_column: String = DEFAULT_RECORD_INDEX_COLUMN, latitude_column: String = DEFAULT_LATITUDE_COLUMN,
                  longitude_column: String = DEFAULT_LONGITUDE_COLUMN, depth_column: String = DEFAULT_DEPTH_COLUMN,
                  height_column: String = DEFAULT_HEIGHT_COLUMN, description_column: String = DEFAULT_DESCRIPTION_COLUMN) -> bool: # returns success
    if not isinstance(file_location, str) or not isinstance(save_directory, str):
        return False

    wb = openpyxl.load_workbook(file_location)
    sheet = wb.active
    if sheet is None:
        return False

    # resolve column indexes from the header row by name
    RECORD_INDEX_COL = find_column(sheet, record_index_column)
    LATITUDE_COL = find_column(sheet, latitude_column)
    LONGITUDE_COL = find_column(sheet, longitude_column)
    DEPTH_COL = find_column(sheet, depth_column)
    HEIGHT_COL = find_column(sheet, height_column)
    FEATURE_DESCRIPTION_COL = find_column(sheet, description_column)

    if None in (RECORD_INDEX_COL, LATITUDE_COL, LONGITUDE_COL, DEPTH_COL, HEIGHT_COL, FEATURE_DESCRIPTION_COL):
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

        depth = float(str(sheet.cell(row=j, column=DEPTH_COL).value).replace(",", "."))

        if len(str(latitude)) > latitude_max_len:
            latitude_max_len = len(str(latitude))
        if len(str(longitude)) > longitude_max_len:
            longitude_max_len = len(str(longitude))

        if (latitude == 0 or longitude == 0):
            record.latitude = 0
            record.longitude = 0
            record.height = 0
        else:
            e = convertToJTSK(latitude, longitude, height)

            if e is None:
                return False

            record.latitude = abs(e[0]) if not negative_xy else -abs(e[0])
            record.longitude = abs(e[1]) if not negative_xy else -abs(e[1])
            if convert_height:
                record.height = round(e[2], 2) if e[2] is not None else str(0)
            else:
                record.height = height if height is not None else str(0)

            if subtract_depth:
                record.height = record.height - depth

        record.description = sheet.cell(row=j, column=FEATURE_DESCRIPTION_COL).value

        records.append(record)

    # sort the records just in case
    sorted(records, key=lambda record: record.index)
    df = pd.DataFrame([(r.index, r.latitude if r.latitude!=0 else str(0), r.longitude if r.longitude!=0 else str(0), r.height if r.height!=0 else str(0), r.description) for r in records]) #, columns=["# Číslo bodu", "Y", "X", "Z", "Popis bodu"])

    file = file_location.split("/")
    file_name = file[-1].split(".")[0]

    with open(save_directory + "/" + file_name + ".txt", "w", encoding="utf-8") as f:

        f.write(df.to_string(index=False, header=False))

    return True
