import tkinter
from tkinter import messagebox, filedialog, ttk

class cookieMonsterFront():

    def select_folder():
        global input_folder
        folder_selected = filedialog.askdirectory()
        input_folder = folder_selected

window = tkinter.Tk()

window.title("Cookie Monster")

def progressBar():
    progressBarWindow = tkinter.Toplevel(window)
    progressBarWindow.title("Progress Bar")
    progressBar = ttk.Progressbar(progressBarWindow, orient="horizontal", mode="indeterminate", length=280)
    progressBar.grid(column=0, row=1, columnspan=2)
    progressBar.start()


def start_sorting():
        
    batch_find_replace_delete_and_remove_chars(input_folder, find_chars, replace_text, delete_chars)
    progressBar
    messagebox.showinfo("Success", "Files have been sorted successfully.")
# window.set
backEnd = cookieMonsterBack()


tkinter.Label(window, text="Select Folder:").grid(row=0, column=0, padx=10, pady=10)
tkinter.Entry(window, text=input_folder, width=50).grid(row=0, column=1, padx=10, pady=10)
tkinter.Button(window, text="Browse", command=select_folder).grid(row=0, column=2, padx=10, pady=10)

tkinter.Label(window, text="Word Count Limit:").grid(row=1, column=0, padx=10, pady=10)
word_limit_entry = tkinter.Entry(window, width=10)
word_limit_entry.grid(row=1, column=1, padx=10, pady=10)

tkinter.Button(window, text="Start Sorting", command=start_sorting).grid(row=2, column=0, columnspan=3, pady=20)

window.mainloop()
