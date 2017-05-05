# !/usr/bin/python
# -*- coding: utf-8 -*-

## Wrap each files to its own folder, run once

# import os
# import shutil
#
# dir_name = "D:\\DOCUMENTS\\Google_Link\\all docx"
# docx_files = os.listdir(dir_name)
#
# for i in docx_files:
#   os.mkdir(os.path.join(dir_name , i.split(".")[0]))
#   shutil.move(os.path.join(dir_name , i), os.path.join(dir_name , i.split(".")[0]))

## This is where the actual code starts 
from docx import Document
import sys, os, re

reload(sys)
sys.setdefaultencoding('utf8')
root_dir = "C:\\Google_Link" # The path of the parent folder which have all the docx files

for root, dirs, files in os.walk(root_dir):
    for name in files:
        doc_file = os.path.join(root, name)
        if doc_file.endswith('docx'):
            main_file = Document(doc_file)
            table = main_file.tables[1]  # I chose to keep this same for all document. You can add another layer to search everywhere

            data = []
            keys = None

            for i, row in enumerate(table.rows):
                text = (cell.text for cell in row.cells)

                if i == 0:
                    keys = tuple(text)
                    continue

                row_data = tuple(text)
                data.append(row_data)

            regexReference = re.compile("(C?..-[0-9-]+)")
            regexCoordinate = re.compile(r'(N-(.{,12})([0-9]|\')|[0-9].{,12}N)[;, ]+(E-(.{,12})([0-9]|\')|[0-9].{,12}E)')

            result = []
            for item in data:
                tmp = dict()
                matchReference = regexReference.search(item[1])
                matchCoordinate = regexCoordinate.search(unicode(item[2]))
                if matchReference:
                    tmp['reference'] = matchReference.group()
                if matchCoordinate:
                    tmp['y'] = matchCoordinate.group(1)
                    tmp['x'] = matchCoordinate.group(4)
                tmp['location'] = unicode(item[3])
                tmp['description'] = unicode(item[2])
                result.append(tmp)

            for rs in result:
                if 'reference' in rs:
                    kml_file_path = os.path.dirname(doc_file)
                    with open(os.path.join(kml_file_path, rs['reference']+'.kml'), 'w') as f:
                       # Writing the kml file.
                       f.write("<?xml version='1.0' encoding='UTF-8'?>\n")
                       f.write("<kml xmlns='http://earth.google.com/kml/2.1'>\n")
                       f.write("<Document>\n")
                       f.write("   <name>" + rs['reference'] + '.kml' + "</name>\n")
                       f.write("   <Placemark>\n")
                       f.write("       <name>" + rs['reference'] + "</name>\n")
                       f.write("       <description>" + "\n \n" + rs['description'] + "</description>\n")
                       f.write("       <Point>\n")
                       f.write("           <coordinates>" + rs['x'] + "," + rs['y']  + "</coordinates>\n")
                       f.write("       </Point>\n")
                       f.write("   </Placemark>\n")
                       f.write("</Document>\n")
                       f.write("</kml>\n")
                       f.close()

#--------------------------------------------------------#

# for rs in extractTableRow():
#     if 'reference' in rs:
        # print "(" + rs['x'] + ")(" + rs['y'] + ")"
        # print rs['y']
        # rx = re.compile(r'((?P<degree>\d+)°\s*(?P<minute>[^\'´]+)[\'´]?\b(?P<second>[^\"´´]+))')        #
        # def convert(match):
        #     try:
        #         degree = float(match.group('degree'))
        #         minute = float(match.group('minute'))
        #         second = float(match.group('second'))
        #         dd = degree + minute / 60
        #     except:
        #         dd = -1
        #     finally:
        #         return dd

        # coordinates_long = [convert(match) for match in rx.finditer(rs['x'])]
        # coordinates_lat = [convert(match) for match in rx.finditer(rs['y'])]
