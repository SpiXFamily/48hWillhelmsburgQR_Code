import tkinter as tk
from tkinter import filedialog
import pandas as pd
import qrcode
from PIL import Image
import xlrd
from openpyxl import load_workbook
import pyshorteners
import re

"""
Description:
This Programm shall collect the colums of the rows of street and house number from a .xls file.
The columns of the each rows shall be combined with eachother to a string.
The string has to be aligned to a map application link where people can scan the code and see the location on the map application.

"""

class App:
    def __init__(self, master):
        self.master = master
        master.title("QR Code Generator")
            #generate the "Select excel file Button"
        self.file_button = tk.Button(master, text="Wähle eine (Excel) Datei aus", command=self.load_data)
        self.file_button.pack()
            #generate the "Generate qrcodes Button" and disable it
        self.generate_button = tk.Button(master, text="Generiere QR Codes", command=self.generate_qr_codes, state=tk.DISABLED)
        self.generate_button.pack()

    def load_data(self):
            # Open a file dialog to select the Excel file
        file_path = filedialog.askopenfilename()
        if file_path.endswith('.xlsx'):
            try:
            # Use openpyxl engine to read .xlsx files
                self.df = pd.read_excel(file_path, engine='openpyxl')
                # Make the Generate Button usable if this function completes
                self.generate_button.config(state=tk.NORMAL)
            except:
                print('Failed to read the xlsx file')
        elif file_path.endswith('.xls'):

            try:
            # Use xlrd engine to read .xls files
                self.df = pd.read_excel(file_path, engine='xlrd')
                # Make the Generate Button usable if this function completes
                self.generate_button.config(state=tk.NORMAL)
            except:
                print('Failed to read the xls file')
        elif file_path.endswith('.xltx'):

            try:
            # Use xlrd engine to read .xltx files
                self.df = pd.read_excel(file_path, engine='openpyxl')
                # Make the Generate Button usable if this function completes
                self.generate_button.config(state=tk.NORMAL)
            except:
                print('Failed to read the xltx file')
        else:
            print('Unsupported file format')



    def generate_qr_codes(self):
            # Iterate over the rows in the DataFrame and generate a QR code for each row
        for index, row in self.df.iterrows():
            #check if the whole row is empty and break the sequence
            if pd.isnull(row['Strasse' or 'Straße']) and pd.isnull(row['Musik']) and pd.isnull(row['Hausnummer']) and pd.isnull(row['Ort']) == True:
                print("Every QR-Code has been successfully generated?")
                break
            musik = row['Musik']
            musik = re.sub('[^a-zA-Z0-9\n\.]', '', musik)

            #check if the house number in the excel file is empty. if it is empty, use an empty string to avoid "nan" as nan will destroy googles search mode
            if pd.isnull(row['Hausnummer']) == True:
                house_number = ""
            else:
                house_number = row['Hausnummer']

            #check if the street in the excel file is empty. If it is empty, use the Location Name of musical event.
            if pd.isnull(row['Strasse' or 'Straße']) == True:
                street = row ['Ort']
            else:
                street = row['Strasse' or 'Straße']

            ext_url = f"{str(street)}+{str(house_number)}+Hamburg&travelmode=walking"
            maps_url = 'https://www.google.com/maps/dir/?api=1&hl=de&destination='
            long_url = str(maps_url)+ (str(ext_url))
            ## Create an instance of the pyshorteners.Shortener class with a bitly API key (((BITLY API HAS A MONTHLY LIMIT, CUT BECAUSE OF MONEY ISSUES)))
            #s = pyshorteners.Shortener(api_key='50c22e4e8b074dd72382e7411c4293e5fbdb2fd6')
            #Shorten the URL using the Bitly API
            #short_url = s.bitly.short(long_url)
                    # Print the shortened URL FOR TESTING PURPOSES
            #print(short_url)

                     # QR Code generating
                    #get an image in the middle of the generated qrcodes
            qr = qrcode.QRCode(version=1, box_size=10, border=5)
            qr.add_data(long_url)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            img.save(f"codes/qrcode_{index+1}_{musik}.png")
            print(musik)
            print(long_url)
                #open the generated qr code
            im = Image.open(f"codes/qrcode_{index+1}_{musik}.png")
                #convert the logo into a format that makes it readable
            im = im.convert("RGBA")

            logo = Image.open('48Logo.png')
            logo = logo.convert(im.mode)

            im_width, im_height = im.size
            logo_width, logo_height = logo.size

            max_size = (im_width * 0.30, im_height * 0.30)
            logo.thumbnail(max_size)

            box_left = (im_width - logo.size[0]) // 2
            box_upper = (im_height - logo.size[1]) // 2
            box_right = box_left + logo.size[0]
            box_lower = box_upper + logo.size[1]

            box = (box_left, box_upper, box_right, box_lower)

            im_crop = im.crop(box)

            im_crop.paste(logo, (0, 0), logo)

            im.paste(im_crop, box)

            im.save(f"logocodes/qrcode_logo_{index+1}_{musik}.png")

root = tk.Tk()
app = App(root)
root.mainloop()
