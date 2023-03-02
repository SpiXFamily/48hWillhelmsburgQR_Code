import pandas as pd
import qrcode
from PIL import Image
df = pd.read_excel('data.xlsx', sheet_name='Sheet1', usecols=['Straße', 'Hausnummer'])




for index, row in df.iterrows():
    street = row['Straße']
    house_number = row['Hausnummer']
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(f"{Straße} {Hausnummer}")
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(f"qrcode_{index}.png")

