#WordWordCounter.py>
import os
import shutil
import win32com.client
import re

def count_words_in_docx(input_folder, wordLimit):
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



            patterns = [',', '-', '=', '𝜃', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '.', '?', ':', ';', '(', ')', '[', ']', '{', '}', '/', '*', '+', '=', '|', '&', '%', '@', '~', '`', '"', '°', '−', '×', '±', '≈', '∆', '>', '<', '>=', '<=', '=', 'ϕ', 'φ', 'Φ', 'Ω', 'Å', '𝜙']            #Had to remove !, \\  and ^

            content = doc.Content

            for pattern in patterns:
                find = content.Find
                find.ClearFormatting()

                find.Text = pattern
                find.Replacement.ClearFormatting()
                find.Replacement.Text = ""
                find.Execute(Replace=2, MatchWildcards=False)
                print(f"Removed {pattern} from {docx_file}")

                    #If not work: try story.Delete()
            
            doc.Save()

            # Count the words in the document
            word_count = doc.ComputeStatistics(0, True)  # 0 for wdStatisticWords

            # Close the document without saving changes
            doc.Close(SaveChanges=False)

            # Create a copy of the document with word count appended to filename

            #above_2000_folder = os.path.join(input_folder, 'Above ' + str(wordLimit)) 
            above_2000_folder = os.path.join(input_folder, "Processed")
            #already_under_2000_folder = os.path.join(input_folder, 'Already under ' + str(wordLimit))
            os.makedirs(above_2000_folder, exist_ok=True)
            #os.makedirs(already_under_2000_folder, exist_ok=True)
            
            if word_count > int(wordLimit):
                output_subfolder = above_2000_folder
            else:
                #output_subfolder = already_under_2000_folder
                output_subfolder = above_2000_folder
            
            new_file_name = f"{word_count}_{docx_file}"
            new_file_path = os.path.join(output_subfolder, new_file_name)

            shutil.copyfile(doc_path, new_file_path)

            print(f"Word count for '{docx_file}': {word_count}. Copied to '{new_file_path}'")

        except Exception as e:
            print(f"Error processing file '{docx_file}': {e}")

    # Quit Word application
    word_app.Quit()

def searchTextBoxes(input_path):
    textBoxText = []
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(input_path)
    try:
        for sh in doc.Shapes:
            if sh.Type == 17:
                print(sh.Name)
                print(sh.TextFrame.TextRange.Text)
                #doc.Range(0,0).InsertBefore(sh.TextFrame.TextRange.Text)
                textBoxText.append(sh.TextFrame.TextRange.Text)
        doc.Save()
    except Exception as e:
        print(f"Error processing file '{input_path}': {e}")
    finally:
        doc.Close()
        word.Quit()
        return textBoxText