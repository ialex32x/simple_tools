#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import sys
import codecs
import urllib

from openpyxl import load_workbook

def parseExcel(file_name):
    workbook = load_workbook(file_name, read_only = True, data_only = True)
    workbook_obj = []

    try:
        for sheet in workbook:
            sheet_obj = {}
            workbook_obj.append(sheet_obj)

            sheet_name = sheet.title
            rows_obj = []
            sheet_obj["sheet"] = sheet_name
            sheet_obj["data"] = rows_obj
            for (row_index, row) in enumerate(sheet.rows):
                row_obj = []
                rows_obj.append(row_obj)

                for cell in row:
                    cell_value = cell.value
                    row_obj.append(cell_value)
    except Error as err:
        print(err)

    return workbook_obj

def export_single(wb, output_filename):
    outfile = codecs.open(output_filename, 'w', 'utf-8') 
    outfile.write(u"""\
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
""")
    for sheetIndex, sheetValue in enumerate(wb):
        data = sheetValue["data"]
        for dataIndex, dataValue in enumerate(data):
            # print(dataIndex)
            outfile.write(u"""\
        <key>%d</key>
        <dict>
""" % (dataIndex + 1))
            for dataValueIndex, dataValueValue in enumerate(dataValue):
                outfile.write(u"""\
            <key>%s</key>
            <string>%s</string>
""" % (chr(dataValueIndex + ord("A")), dataValueValue))
            outfile.write(u"""\
        </dict>
""")
    outfile.write(u"""\
</dict>
</plist>
""")
    outfile.close()    

argc = len(sys.argv)
plist = None
src = None

if argc == 2:
    src = sys.argv[1]
    basename = os.path.basename(src)
    extensionindex = basename.rfind(".")
    simplename = basename if extensionindex < 0 else basename[:extensionindex]
    plist = simplename + ".plist"
elif argc == 3:
    src = sys.argv[1]
    plist = sys.argv[2]
else:
    print("invalid option")
    print("usage: %s input_file [output_file]" % (os.path.basename(sys.argv[0])))
    exit(1)

wb = parseExcel(src)
export_single(wb, plist)
