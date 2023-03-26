import tkinter as tk
import pandas as pd
import qrcode
from tkinter import Label, filedialog, BooleanVar
from PIL import Image, ImageTk
import xlrd
from openpyxl import load_workbook
import pyshorteners
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
class App:
    def __init__(self, master):
        self.master = master
        master.title("QR Code Generator")
        master.geometry("900x600")

            #create a button to exit the app
        self.kill_button = tk.Button(master, text="Schließe die App", command=master.destroy)
        self.kill_button.pack(side="left", anchor="nw", fill=tk.X)

            #create a button to open the folder with the generated qr-codes
        self.explorer_button = tk.Button(master, text="Öffne den Ordner mit den QR-Codes", command=self.open_explorer, state=tk.DISABLED)
        self.explorer_button.pack(side="right", anchor="ne")

            #generate the "Select excel file Button"
        self.file_button = tk.Button(master, text="Wähle eine (Excel) Datei aus", command=self.load_data)
        self.file_button.pack(side="top",anchor="n",padx=200)

            #generate the "Generate qrcodes Button" and disable it
        self.generate_button = tk.Button(master, text="Generiere QR Codes", command=function.generate_qr_codes, state=tk.DISABLED)
        self.generate_button.pack(side="bottom",anchor="s")

            #label for the generate button
        self.generate_label = tk.Label(master, text="Bitte zuerst eine Datei auswählen")
        self.generate_label.pack(side="bottom",anchor="s")

            #label for the size slider
        self.label_size_slider = tk.Label(master, text="Wähle die QR-Code Größe. 10 ist Standard.")
        self.label_size_slider.pack(side="top", anchor="n")

            #create a slider for the qrcode size
        self.size_slider = tk.Scale(master, from_=1, to=20, orient='horizontal')
        self.size_slider.pack(side="top",anchor="n")
        self.size_slider.set(10)

            #label for the border slider
        self.label_border_slider = tk.Label(master, text="Wähle die QR-Code Randgröße. 5 ist Standard.")
        self.label_border_slider.pack(side="top", anchor="n")

            #create a slider for the qrcode border size
        self.border_slider = tk.Scale(master, from_=1, to=10, orient='horizontal')
        self.border_slider.pack(side="top",anchor="n")
        self.border_slider.set(5)

            #create labels for User feedback
        self.correct_file = tk.Label(master, text="Richtige Datei ausgewählt!")
        self.incorrect_file = tk.Label(master, text="Falsche Datei ausgewählt :/ bitte wähle eine .xls .xlsx oder .xltx Datei!")
        self.number_qrcodes = tk.Label(master)

    def open_explorer():
        directory = filedialog.askopenfilename(initialdir="./logocodes")

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
                # Use xlrd engine to read .xltx files

                self.df = pd.read_excel(file_path, engine='openpyxl')

                # Make the Generate Button usable if this function completes
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

root = tk.Tk()
app = App(root)
root.mainloop()
