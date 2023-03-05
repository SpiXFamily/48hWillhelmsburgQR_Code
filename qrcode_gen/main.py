import pandas as pd
import qrcode
df = pd.read_excel('data.xls') 
"""
Description:
This Programm shall collect the colums of the rows of street and house number from a .xls file.
The columns of the each rows shall concatinated with eachother to a string. 
The string has to be aligned to a map application link where people can scan the code and see the location on the map application.

"""
for index, row in df.iterrows():
    # find the rows of street and house number
    house_number = row['Hausnummer']
    #house_number_str = str(house_number)
    street = row['Strasse' or 'Straße']
    plz = row['PLZ']
    int_plz = int(plz)
    # TODO concat street with house_number
    url = str(street) + str(house_number)
    # Map link
    ext_url = f"{str(street)}+{str(house_number)},+{str(int_plz)}+Hamburg"
    maps_url = 'https://google.com/maps/place/'
    # QR Code generating
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(f"{maps_url}{ext_url}") 
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(f"codes/qrcode_{index}.png")
