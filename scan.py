#! python3
# Drew McLaughlin

import cv2
import numpy as np
import pandas as pd
import pandas.io.formats.excel
import pyperclip
from pyzbar.pyzbar import decode
import sys

# The QR code below was copied from this repository: https://github.com/pruthvi03/Barcode-And-Qrcode-Scanner
def decoder(image, data_row):
    gray_img = cv2.cvtColor(image,0)
    barcode = decode(gray_img)

    for obj in barcode:
        points = obj.polygon
        (x,y,w,h) = obj.rect
        pts = np.array(points, np.int32)
        pts = pts.reshape((-1, 1, 2))
        cv2.polylines(image, [pts], True, (0, 255, 0), 3)

        barcodeData = obj.data.decode("utf-8")
        barcodeType = obj.type
        string = "Data: " + str(barcodeData) + " | Type: " + str(barcodeType)
        
        cv2.putText(frame, string, (x,y), cv2.FONT_HERSHEY_SIMPLEX,0.8,(0,0,255), 2)
        
        data_row.append(barcodeData)
        print("Barcode: " + barcodeData)

data_row = []

try:
    # Set the Column Header
    pack_name = str(input("Enter the PTCGO Name: "))

    # Set the Filename
    file_name = str(input("Enter the output filename: "))

    
    cap = cv2.VideoCapture(0)
    while True:
        ret, frame = cap.read()
        decoder(frame, data_row)
        cv2.imshow('Image', frame)
        code = cv2.waitKey(100)
        if code == ord('q'):
            break
except KeyboardInterrupt:
    # Remove duplicates from the list
    data_row = list(dict.fromkeys(data_row))

    # Create the dataframe
    df = pd.DataFrame({pack_name:data_row})
    
    # Remove the pandas header format
    pandas.io.formats.excel.ExcelFormatter.header_style = None

    # Create the file
    writer = pd.ExcelWriter("scanner\\"+file_name + ".xlsx")
    df.to_excel(writer, sheet_name="Sheet1", index=False)

    # Automatically Set the Column Width
    for column in df:
        column_width = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
    
    # Save the file
    writer.save()

    # Copy the list to the clipboard as a string
    data_string = '\n'.join(data_row)
    pyperclip.copy(data_string)

    # Exit the program
    sys.exit()
