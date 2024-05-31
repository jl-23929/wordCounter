from tkinter import filedialog
from tkinter import *
import customtkinter
from PIL import Image

window = customtkinter.CTk()

window.geometry("700x400")

window.title("Cookie Monster")
window.resizable(0,0)

def selectFolder():
    global input_folder
    folder_selected = filedialog.askdirectory()
    input_folder = folder_selected

column1 = 0.1
column2 = 0.2
column3 = 0.35

image_path = r"C:\Users\james680384\Pictures\Picture1.png"

pil_image = Image.open(image_path)

image = customtkinter.CTkImage(light_image=pil_image, dark_image=pil_image, size=(200,200))
imageLabel = customtkinter.CTkLabel(window, image=image, text="")
imageLabel.place(relx = 0.25, rely = 0.7, anchor=CENTER)


bold = customtkinter.CTkFont(family="Arial Black", size=32)

infoHeading = customtkinter.CTkLabel(window, text="Cookie Monster", font=bold)
infoHeading.place(relx=0.5, rely=column1, anchor=CENTER)

infoLabel = customtkinter.CTkLabel(window, text="To use Cookie Monster, ..................")
infoLabel.place(relx=0.5, rely=column2, anchor=CENTER)

selectFolderLabel = customtkinter.CTkLabel(window, text="Select Folder:")
selectFolderLabel.place(relx=0.2, rely = column3, anchor=CENTER)

folderEntry = customtkinter.CTkEntry(window, placeholder_text="Enter a path or browse", width=160)
folderEntry.place(relx = 0.5, rely = column3, anchor=CENTER)

browseButton = customtkinter.CTkButton(master=window, text="Browse", command=selectFolder)
browseButton.place(relx=0.8, rely=column3, anchor=CENTER)

wordLimitEntry = customtkinter.CTkEntry(window, placeholder_text="Word Limit")
wordLimitEntry.place(relx=0.5, rely = 0.5, anchor=CENTER)


sortButton = customtkinter.CTkButton(master=window, text="Start Sorting")
sortButton.place(relx=0.5, rely=0.75, anchor = CENTER)

window.mainloop()