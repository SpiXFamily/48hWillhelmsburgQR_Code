import tkinter as tk
from tkinter import filedialog
import pandas as pd
import qrcode
from PIL import Image

class App:
    def __init__(self, master):
        self.master = master
        master.title("QR Code Generator")

        self.load_button = tk.Button(master, text="Load Data", command=self.load_data)
        self.load_button.pack()

        self.generate_button = tk.Button(master, text="Generate QR Codes", command=self.generate_qr_codes)
        self.generate_button.pack()

    def load_data(self):
        # Open a file dialog to select the Excel file
        file_path = filedialog.askopenfilename()
        
        # Load the data into a Pandas DataFrame
        self.df = pd.read_excel(file_path)

    def generate_qr_codes(self):
        # Iterate over the rows in the DataFrame and generate a QR code for each row
        for index, row in self.df.iterrows():
            musik = row['Musik']
            musik = str(musik).replace("/", "")

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

            url = str(street) + str(house_number)
            ext_url = f"{str(street)}+{str(house_number)}+Hamburg&travelmode=walking"
            maps_url = 'https://www.google.com/maps/dir/?api=1&hl=de&destination='

            qr = qrcode.QRCode(version=1, box_size=10, border=5)
            qr.add_data(f"{maps_url}{ext_url}") 
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            img.save(f"codes/qrcode_{index+1}_{musik}.png")
            print(f"{maps_url}{ext_url}")

            im = Image.open(f"codes/qrcode_{index+1}_{musik}.png")
            im = im.convert("RGBA")

            logo = Image.open('48Logo.png')
            logo = logo.convert(im.mode)

            im_width, im_height = im.size
            logo_width, logo_height = logo.size

            max_size = (im_width * 0.40, im_height * 0.40)
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
