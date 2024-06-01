#WordWordCounter.py>
import os
import shutil
import win32com.client

def count_words_in_docx(input_folder):
    # Initialize Word application
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False  # Hide Word application window

    # Get all docx files in the specified folder
    docx_files = [file for file in os.listdir(input_folder) if file.endswith('.docx')]

    for docx_file in docx_files:
        try:
            # Open the Word document
            doc_path = os.path.join(input_folder, docx_file)
            doc = word_app.Documents.Open(doc_path)

            # Count the words in the document
            word_count = doc.ComputeStatistics(0)  # 0 for wdStatisticWords

            # Close the document without saving changes
            doc.Close(SaveChanges=False)

            # Create a copy of the document with word count appended to filename
            new_file_name = f"{word_count}_{docx_file}"
            new_file_path = os.path.join(input_folder, new_file_name)
            shutil.copyfile(doc_path, new_file_path)

            print(f"Word count for '{docx_file}': {word_count}. Copied to '{new_file_path}'")

        except Exception as e:
            print(f"Error processing file '{docx_file}': {e}")

    # Quit Word application
    word_app.Quit()