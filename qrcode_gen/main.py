import pandas as pd
import qrcode
import tkinter
from PIL import Image
df = pd.read_excel('data.xls')

"""
Description:
This Programm shall collect the colums of the rows of street and house number from a .xls file.
The columns of the each rows shall be combined with eachother to a string. 
The string has to be aligned to a map application link where people can scan the code and see the location on the map application.

"""

for index, row in df.iterrows():
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
    
    #plz = row['PLZ']
    #int_plz = int(plz)
    # TODO concat street with house_number
    url = str(street) + str(house_number)
    # Map links
    ext_url = f"{str(street)}+{str(house_number)}+Hamburg&travelmode=walking" #+{str(int_plz)}
    maps_url = 'https://www.google.com/maps/dir/?api=1&hl=de&destination='
    
    # QR Code generating
    #get an image in the middle of the generated qrcodes


    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(f"{maps_url}{ext_url}") 
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(f"codes/qrcode_{index+1}_{musik}.png")
    print(f"{maps_url}{ext_url}")
    
    # Load the image to be altered
    im = Image.open(f"codes/qrcode_{index+1}_{musik}.png")
    im = im.convert("RGBA")

    # Load the logo image and convert it to the same mode as the QR code image
    logo = Image.open('48Logo.png')
    logo = logo.convert(im.mode)

        # Calculate the coordinates for the box where the logo will be pasted
    im_width, im_height = im.size
    logo_width, logo_height = logo.size

        # Calculate the maximum size of the logo to fit within the dimensions of the QR code
    max_size = (im_width * 0.40, im_height * 0.40)
    logo.thumbnail(max_size)

    box_left = (im_width - logo.size[0]) // 2
    box_upper = (im_height - logo.size[1]) // 2
    box_right = box_left + logo.size[0]
    box_lower = box_upper + logo.size[1]

    box = (box_left, box_upper, box_right, box_lower)

        # Crop the region of the image where the logo will be pasted
    im_crop = im.crop(box)

        # Paste the resized logo onto the cropped region
    im_crop.paste(logo, (0, 0), logo)

        # Replace the original region with the modified region
    im.paste(im_crop, box)

        # Display the final image
    im.save(f"logocodes/qrcode_logo_{index+1}_{musik}.png")

