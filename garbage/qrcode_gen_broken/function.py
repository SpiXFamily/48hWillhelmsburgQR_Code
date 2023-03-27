import qrcode
from PIL import Image, ImageTk
import re
import xlrd
import pandas as pd

def generate_qr_codes():
    global counter
    global size_slider
            #Iterate over the rows in the DataFrame and generate a QR code for each row
    for index, row in df.iterrows():

            #check if the whole row is empty with regex and break the sequence
        if pd.isnull(row['Strasse' or 'Straße']) and pd.isnull(row['Musik']) and pd.isnull(row['Hausnummer']) and pd.isnull(row['Ort']) == True:
            print("Every QR-Code has been successfully generated?")
            number_qrcodes.pack_forget()
            number_qrcodes = Label(root, text=str(counter) + " QR-Codes wurden erfolgreich generiert!")
            number_qrcodes.pack(side="bottom")
            counter = 0
            explorer_button.config(state=tk.NORMAL)
            break
            #count the amount of times the sequence has been run through to display the amount of qr-codes generated

        counter = counter +1
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

        # QR Code generating
        #get an image in the middle of the generated qrcodes

        qr = qrcode.QRCode(version=1, box_size=size_slider.get(), border=border_slider.get())
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
        #open the logo
        logo = Image.open('48hLogoTransparent.png')
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
        street2 = "/" + str(street)
        print(street2)
        im.save(f"logocodes/byname/qrcode_logo_{index+1}_{musik}.png")
        im.save(f"logocodes/bylocation/location_{street}{house_number}.png")
        # im.save(f"logocodes_sorted{street2}_{index+1}_{musik}.png")

        #show the qrcodes on the bottom of the Ui for a cool effect
        image = Image.open (f"logocodes/byname/qrcode_logo_{index+1}_{musik}.png")
        photo = ImageTk.PhotoImage(image)
        label = Label(root, image=photo)
        label.pack(side="bottom")
        root.update()
        label.destroy()
