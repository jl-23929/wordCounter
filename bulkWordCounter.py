from docx import Document
import os
import shutil
import sys

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
    # Verify if the folder path exists
    if not os.path.isdir(folder_path):
        print(f"The folder path '{folder_path}' does not exist.")
        sys.exit(1)
    
    # Create folders for sorting
    more_words_folder = os.path.join(folder_path, "Over Word Count")
    less_words_folder = os.path.join(folder_path, "Under Word Count")
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

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python sort_files.py <path_to_folder> <word_count_limit>")
        sys.exit(1)
    
    folder_path = sys.argv[1]
    try:
        word_limit = int(sys.argv[2])
    except ValueError:
        print("The word count limit must be an integer.")
        sys.exit(1)
    
    print(f"Folder path: {folder_path}")
    print(f"Word count limit: {word_limit}")
    sort_files_by_word_count(folder_path, word_limit)
