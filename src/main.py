import os # To create files
# for GUI creation
import tkinter as tk
from tkinter import Label, filedialog, BooleanVar
# for working with Excel files
import pandas as pd
import xlrd
# for generating QR_Code
import qrcode
from PIL import Image, ImageTk
import re # for regex
import requests # for generating urls
import shutil

# For optimizing QR_Code size
#import pyshorteners

# create folders if not excists
folder_path = "logocodes"
folder_path1 = "codes"
folder_path2 = "logo"
# create the folder and any necessary parent folders
os.makedirs(folder_path, exist_ok=True)
os.makedirs(folder_path1, exist_ok=True)
os.makedirs(folder_path2, exist_ok=True)
# create an empty folder inside the newly created folder
subfolder_path = os.path.join(folder_path, "byname")
os.makedirs(subfolder_path, exist_ok=True)
subfolder_path = os.path.join(folder_path, "bylocation")
os.makedirs(subfolder_path, exist_ok=True)

# Applation user interface
counter = 0
class App:
    global logo_was_selected
    logo_was_selected = False

    def __init__(self, master):
        self.master = master
        master.title("QR Code Generator")
        master.geometry("1200x800")

        # create a button to exit the app
        self.kill_button = tk.Button(master, text="Schließe die App", command=master.destroy)
        self.kill_button.pack(side="left", anchor="nw", fill=tk.X)

        # create a button to open the folder with the generated qr-codes with logo
        self.codes_logo_button = tk.Button(master, text="Öffne den Ordner mit den QR-Codes mit Logo!", command=self.open_codes_logo)
        self.codes_logo_button.pack(side="right", anchor="ne")

        # create a button to open the folder with the generated qr-codes
        self.codes_button = tk.Button(master, text="Öffne den Ordner mit den QR-Codes", command=self.open_codes)
        self.codes_button.pack(side="right", anchor="ne")

        # generate the "Select logo png Button"
        self.logo_button = tk.Button(master, text="Wähle ein Logo für deine qrcodes aus (optional)", command=self.select_logo)
        self.logo_button.pack(side="top",anchor="n",padx=150)

        # generate the "Select excel file Button"
        self.file_button = tk.Button(master, text="Wähle eine Excel (.xls, .xlsx, .xltx) Datei aus", command=self.load_data)
        self.file_button.pack(side="top",anchor="n")

        # generate the "Generate qrcodes Button" and disable it
        self.generate_button = tk.Button(master, text="Generiere QR Codes", command=self.generate_qr_codes, state=tk.DISABLED)
        self.generate_button.pack(side="bottom",anchor="s")

        # label for the generate button
        self.generate_label = tk.Label(master, text="Bitte zuerst eine (Excel) Datei auswählen")
        self.generate_label.pack(side="bottom",anchor="s")

        #label for no selected logo user feedback
        self.logo_not_selected_label = tk.Label(master, text="Es wurde kein Logo ausgewählt!")
        self.logo_not_selected_label.pack(side="bottom",anchor="s")

        # label for the size slider
        self.label_size_slider = tk.Label(master, text="Wähle die QR-Code Größe. 10 ist Standard.")
        self.label_size_slider.pack(side="top", anchor="n")

        # create a slider for the qrcode size
        self.size_slider = tk.Scale(master, from_=1, to=15, orient='horizontal')
        self.size_slider.pack(side="top",anchor="n")
        self.size_slider.set(10)

        # label for the border slider
        self.label_border_slider = tk.Label(master, text="Wähle die QR-Code Randgröße. 5 ist Standard.")
        self.label_border_slider.pack(side="top", anchor="n")

        # create a slider for the qrcode border size
        self.border_slider = tk.Scale(master, from_=1, to=10, orient='horizontal')
        self.border_slider.pack(side="top",anchor="n")
        self.border_slider.set(5)

        # label for the size slider
        self.logo_size_slider = tk.Label(master, text="Wähle die Logo Größe. (optional) 0.30 ist Standard.")
        self.logo_size_slider.pack(side="top", anchor="n")

        # create a slider for the qrcode size
        self.logo_slider = tk.Scale(master, from_=0.1, to=0.3, resolution=0.01, orient='horizontal')
        self.logo_slider.pack(side="top",anchor="n")
        self.logo_slider.set(0.3)

        # create labels for User feedback
        self.correct_file = tk.Label(master, text="Richtige Datei ausgewählt!")
        self.incorrect_file = tk.Label(master, text="Falsche Datei ausgewählt :/ bitte wähle eine .xls .xlsx oder .xltx Datei!")
        self.logo_selected_label = tk.Label(master, text="Ein Logo wurde ausgewählt!")
        self.number_qrcodes = tk.Label(master)


    def select_logo(self):
        logo_selected = filedialog.askopenfilename(initialdir="./", filetypes=[("PNG files", "*.png")])

        if logo_selected:
            global logo_was_selected
            # Copy the selected file to the "logo" directory with a new name
            logo_name = os.path.basename(logo_selected)  # Get the original file name
            new_logo_name = "ITECHSCHULPROJEKT258.png"  # Set the new file name
            logo_dest = os.path.join("logo", new_logo_name)  # Set the destination path
            shutil.copy2(logo_selected, logo_dest)  # Copy the file to the new location
            # Rename the copied file to the new file name
            os.rename(logo_dest, os.path.join("logo", new_logo_name))
            print(f"Logo copied and saved as {new_logo_name}")
            logo_was_selected = True
            self.logo_not_selected_label.pack_forget()
            self.logo_selected_label.pack(side="bottom",anchor="s")

    def open_codes_logo(self):
        directory = filedialog.askdirectory(initialdir="./logocodes")

    def open_codes(self):
        directory = filedialog.askdirectory(initialdir="./codes")

    def load_data(self):
            # Open a file dialog to select the Excel file
        file_path = filedialog.askopenfilename(initialdir="./")

        if file_path.endswith('.xlsx'):
            try:
                # Use openpyxl engine to read .xlsx files
                self.df = pd.read_excel(file_path, engine='openpyxl')
                # Make the Generate Button usable if this function completes
                self.generate_button.config(state=tk.NORMAL)
                #forget a wrong label and show the correct one
                self.incorrect_file.pack_forget()
                self.generate_label.pack_forget()
                self.correct_file.pack(side="bottom", anchor="s")

            except:
                print('Failed to read the xlsx file')

        elif file_path.endswith('.xls'):
            try:
                # Use xlrd engine to read .xls files
                self.df = pd.read_excel(file_path, engine='xlrd')
                # Make the Generate Button usable if this function completes
                self.generate_button.config(state=tk.NORMAL)
                #forget a wrong label and show the correct one
                self.incorrect_file.pack_forget()
                self.generate_label.pack_forget()
                self.correct_file.pack(side="bottom", anchor="s")
            except:
                print('Failed to read the xls file')

        elif file_path.endswith('.xltx'):

            try:
                # use xlrd engine to read .xltx files
                self.df = pd.read_excel(file_path, engine='openpyxl')
                # make the Generate Button usable if this function completes
                self.generate_button.config(state=tk.NORMAL)
                # forget a wrong label and show the correct one
                self.generate_label.pack_forget()
                self.incorrect_file.pack_forget()
                self.correct_file.pack(side="bottom", anchor="s")
            except:
                print('Failed to read the xltx file')

        else:
            print('Unsupported file format')
            # forget a wrong label and show the correct one
            self.generate_button.config(state=tk.DISABLED)
            self.generate_label.pack_forget()
            self.correct_file.pack_forget()
            self.incorrect_file.pack(side="bottom", anchor="s")

    # TODO outsource for optimisation
    def generate_qr_codes(self):
        global counter
        global size_slider
        global logo_was_selected

        # Iterate over the rows in the DataFrame and generate a QR code for each row
        for index, row in self.df.iterrows():

            # check if the whole row is empty with regex and break the sequence
            if pd.isnull(row['Strasse' or 'Straße']) and pd.isnull(row['Musik']) and pd.isnull(row['Hausnummer']) and pd.isnull(row['Ort']) == True:
                print("Every QR-Code has been successfully generated?")
                self.number_qrcodes.pack_forget()
                self.number_qrcodes = Label(root, text=str(counter) + " QR-Codes wurden erfolgreich generiert!")
                self.number_qrcodes.pack(side="bottom")
                counter = 0
                logo_was_selected = False
                break
            # count the amount of times the sequence has been run through to display the amount of qr-codes generated
            counter = counter +1
                #remove
            musik = row['Musik']
            musik = re.sub('[^a-zA-Z0-9\n\.]', '', musik)

            # check if the house number in the excel file is empty. if it is empty, use an empty string to avoid "nan" as nan will destroy googles search mode
            if pd.isnull(row['Hausnummer']) == True:
                house_number = ""
            else:
                house_number = row['Hausnummer']

            # check if the street in the excel file is empty. If it is empty, use the Location Name of musical event.
            if pd.isnull(row['Strasse' or 'Straße']) == True:
                street = row ['Ort']
            else:
                street = row['Strasse' or 'Straße']

            ext_url = f"{str(street)}+{str(house_number)}+Hamburg&travelmode=walking"
            maps_url = 'https://www.google.com/maps/dir/?api=1&hl=de&destination='
            long_url = str(maps_url)+ (str(ext_url))
            ## Create an instance of the pyshorteners.Shortener class with a bitly API key (((BITLY API HAS A MONTHLY LIMIT, CUT BECAUSE OF MONEY ISSUES)))
            #Shorten the URL using the Bitly API
            #short_url = s.bitly.short(long_url)
            # Print the shortened URL FOR TESTING PURPOSES
            #print(short_url)

            # QR Code generating
            # get an image in the middle of the generated qrcodes
            qr = qrcode.QRCode(version=1, box_size=self.size_slider.get(), border=self.border_slider.get())
            qr.add_data(long_url)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            img.save(f"codes/qrcode_{index+1}_{musik}.png")
            print(musik)
            print(long_url)
            # open the generated qr code
            im = Image.open(f"codes/qrcode_{index+1}_{musik}.png")
            # convert the logo into a format that makes it readable
            im = im.convert("RGBA")
            # TODO
            # open the logo
            if logo_was_selected == True:
                logo = Image.open(f"logo/ITECHSCHULPROJEKT258.png")
                logo = logo.convert(im.mode)

                im_width, im_height = im.size
                logo_width, logo_height = logo.size

                max_size = (im_width * self.logo_slider.get(), im_height * self.logo_slider.get())
                logo.thumbnail(max_size)

                box_left = (im_width - logo.size[0]) // 2
                box_upper = (im_height - logo.size[1]) // 2
                box_right = box_left + logo.size[0]
                box_lower = box_upper + logo.size[1]

                box = (box_left, box_upper, box_right, box_lower)

                im_crop = im.crop(box)

                im_crop.paste(logo, (0, 0), logo)

                im.paste(im_crop, box)

                im.save(f"logocodes/byname/qrcode_logo_{index+1}_{musik}.png")

                ort = row["Ort"]
                if pd.isnull(row['Hausnummer']) and pd.isnull(row['Strasse' or 'Straße']) == True:
                    im.save(f"logocodes/bylocation/location_{ort}.png")
                else:
                    im.save(f"logocodes/bylocation/location_{street}{house_number}.png")

                    # show the qrcodes on the bottom of the Ui for a cool effect
                image = Image.open (f"logocodes/byname/qrcode_logo_{index+1}_{musik}.png")
                photo = ImageTk.PhotoImage(image)
                label = Label(root, image=photo)
                label.pack(side="bottom")
                root.update()
                label.destroy()
            else:
                image = Image.open (f"codes/qrcode_{index+1}_{musik}.png")
                photo = ImageTk.PhotoImage(image)
                label = Label(root, image=photo)
                label.pack(side="bottom")
                root.update()
                label.destroy()
root = tk.Tk()
app = App(root)
root.mainloop()
