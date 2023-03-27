import tkinter as tk
import pandas as pd
import qrcode
from tkinter import Label, filedialog, BooleanVar
from PIL import Image, ImageTk
import xlrd
from openpyxl import load_workbook
import re
import requests
from function import generate_qr_codes


"""
Description:
This Programm shall collect the colums of the rows of street and house number from a .xls file.
The columns of the each rows shall be combined with eachother to a string.
The string has to be aligned to a map application link where people can scan the code and see the location on the map application.

"""
counter = 0
title("QR Code Generator")
geometry("900x600")

#create a button to exit the app
kill_button = tk.Button(text="Schließe die App", command=root.quit())
kill_button.pack(side="left", anchor="nw", fill=tk.X)

#create a button to open the folder with the generated qr-codes
explorer_button = tk.Button(text="Öffne den Ordner mit den QR-Codes", command=open_explorer, state=tk.DISABLED)
explorer_button.pack(side="right", anchor="ne")

#generate the "Select excel file Button"
file_button = tk.Button(text="Wähle eine (Excel) Datei aus", command=load_data)
file_button.pack(side="top",anchor="n",padx=200)

#generate the "Generate qrcodes Button" and disable it
generate_button = tk.Button( text="Generiere QR Codes", command=generate_qr_codes, state=tk.DISABLED)
generate_button.pack(side="bottom",anchor="s")

#label for the generate button
generate_label = tk.Label(text="Bitte zuerst eine Datei auswählen")
generate_label.pack(side="bottom",anchor="s")

#label for the size slider
label_size_slider = tk.Label(text="Wähle die QR-Code Größe. 10 ist Standard.")
label_size_slider.pack(side="top", anchor="n")

#create a slider for the qrcode size
size_slider = tk.Scale(from_=1, to=20, orient='horizontal')
size_slider.pack(side="top",anchor="n")
size_slider.set(10)

#label for the border slider
label_border_slider = tk.Label(text="Wähle die QR-Code Randgröße. 5 ist Standard.")
label_border_slider.pack(side="top", anchor="n")

#create a slider for the qrcode border size
border_slider = tk.Scale(from_=1, to=10, orient='horizontal')
border_slider.pack(side="top",anchor="n")
border_slider.set(5)

#create labels for User feedback
correct_file = tk.Label(text="Richtige Datei ausgewählt!")
incorrect_file = tk.Label(text="Falsche Datei ausgewählt :/ bitte wähle eine .xls .xlsx oder .xltx Datei!")
number_qrcodes = tk.Label()

def open_explorer():
    directory = filedialog.askopenfilename(initialdir="./logocodes")

def load_data():
    # Open a file dialog to select the Excel file
    file_path = filedialog.askopenfilename(initialdir="./")

    if file_path.endswith('.xlsx'):
        try:
            # Use openpyxl engine to read .xlsx files
            df = pd.read_excel(file_path, engine='openpyxl')
            # Make the Generate Button usable if this function completes
            generate_button.config(state=tk.NORMAL)
            #forget a wrong label and show the correct one
            incorrect_file.pack_forget()
            generate_label.pack_forget()
            correct_file.pack(side="bottom", anchor="s")

        except:
            print('Failed to read the xlsx file')

    elif file_path.endswith('.xls'):
        try:
            # Use xlrd engine to read .xls files
            df = pd.read_excel(file_path, engine='xlrd')
            # Make the Generate Button usable if this function completes
            generate_button.config(state=tk.NORMAL)
            # forget a wrong label and show the correct one
            incorrect_file.pack_forget()
            generate_label.pack_forget()
            correct_file.pack(side="bottom", anchor="s")
        except:
            print('Failed to read the xls file')

    elif file_path.endswith('.xltx'):

        try:
            # Use xlrd engine to read .xltx files

            df = pd.read_excel(file_path, engine='openpyxl')

            # Make the Generate Button usable if this function completes
            generate_button.config(state=tk.NORMAL)

            # forget a wrong label and show the correct one
            generate_label.pack_forget()
            incorrect_file.pack_forget()
            correct_file.pack(side="bottom", anchor="s")
        except:
            print('Failed to read the xltx file')

    else:
        print('Unsupported file format')
        # forget a wrong label and show the correct one

generate_button.config(state=tk.DISABLED)
generate_label.pack_forget()
correct_file.pack_forget()
incorrect_file.pack(side="bottom", anchor="s")

root = tk.Tk()
app = App()
root.mainloop()
