from docx import Document
import sys

def count_words(doc_path):
    # Open the Word document
    doc = Document(doc_path)
    
    # Initialize the word count
    word_count = 0
    
    # Iterate through paragraphs in the document
    for para in doc.paragraphs:
        word_count += len(para.text.split())
    
    return word_count

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python count_words.py <path_to_docx_file>")
        sys.exit(1)
    
    doc_path = sys.argv[1]
    try:
        word_count = count_words(doc_path)
        print(f"Word count: {word_count}")
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)
