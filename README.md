# PTCGO-Card-Scanner
This python script will capture the QR code on the PTCGO cards and add them to your clipboard as well as export them to excel.

## Short Guide
1. Install Python 3.9+
2. Install the modules in requirements.txt via pip.  Ex: `pip install pandas`
3. Connect a webcam to your computer.
4. Run the program with `python3 scan.py`
5. Enter the PTCGO Name.  Ex: `Sword & Shield - Chilling Reign` (This is used for the column header)
6. Enter the output filename.  Ex: `Chilling Reign Codes` (This is going to be your excel file.  **Do not add the .xlsx extension.**
7. Once the camera loads (it generally takes about ten seconds after entering the output filename), point the camera at the QR codes on the PTCGO cards.  It is okay if it captures the barcode more than once as the script will remove duplicates at the end.
8. When you are done scanning your cards, press `CTRL + C` to exit the camera function.
9. The program will run for a couple more seconds.  During this time it will create an excel file and copy the list of codes to your clipboard.

## Credits
This program is based off of the [following GitHub repository](https://github.com/pruthvi03/Barcode-And-Qrcode-Scanner).  I modified the script to capture the QR codes and send them to an excel file and the computer's clipboard.
