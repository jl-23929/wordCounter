from docx import Document
import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

def count_words_in_docx(doc_path):
    try:
        # Open the Word document
        doc = Document(doc_path)
        
        # Initialize the word count
        word_count = 0
        
        # Iterate through paragraphs in the document
        for para in doc.paragraphs:
            word_count += len(para.text.split())
        
        return word_count
    except Exception as e:
        print(f"Error processing document '{doc_path}': {e}")
        return None

def sort_files_by_word_count(folder_path, word_limit):
    # Create folders for sorting
    more_words_folder = os.path.join(folder_path, "more_words")
    less_words_folder = os.path.join(folder_path, "less_words")
    os.makedirs(more_words_folder, exist_ok=True)
    os.makedirs(less_words_folder, exist_ok=True)
    
    # Iterate through all files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".docx"):
            doc_path = os.path.join(folder_path, filename)
            print(f"Processing file: {doc_path}")
            word_count = count_words_in_docx(doc_path)
            if word_count is not None:
                if word_count >= word_limit:
                    new_path = os.path.join(more_words_folder, filename)
                else:
                    new_path = os.path.join(less_words_folder, filename)
                
                try:
                    # Move the file to the appropriate folder
                    shutil.move(doc_path, new_path)
                    print(f"Moved '{filename}' to '{new_path}'")
                except Exception as e:
                    print(f"Error moving file '{filename}': {e}")

def select_folder():
    folder_selected = filedialog.askdirectory()
    folder_path.set(folder_selected)

def start_sorting():
    folder = folder_path.get()
    try:
        word_limit = int(word_limit_entry.get())
    except ValueError:
        messagebox.showerror("Invalid Input", "The word count limit must be an integer.")
        return
    
    if not folder:
        messagebox.showerror("Invalid Input", "Please select a folder.")
        return

    sort_files_by_word_count(folder, word_limit)
    messagebox.showinfo("Success", "Files have been sorted successfully.")

# Set up the GUI
root = tk.Tk()
root.title("Word Document Sorter")

folder_path = tk.StringVar()

# Folder selection
tk.Label(root, text="Select Folder:").grid(row=0, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=folder_path, width=50).grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_folder).grid(row=0, column=2, padx=10, pady=10)

# Word count limit input
tk.Label(root, text="Word Count Limit:").grid(row=1, column=0, padx=10, pady=10)
word_limit_entry = tk.Entry(root, width=10)
word_limit_entry.grid(row=1, column=1, padx=10, pady=10)

# Start button
tk.Button(root, text="Start Sorting", command=start_sorting).grid(row=2, column=0, columnspan=3, pady=20)

root.mainloop()
