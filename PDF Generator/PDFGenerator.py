from docx import Document
import random
import string

def generate_word_document(file_name):
    doc = Document()
    num_paragraphs = random.randint(1, 15)  # Generate a random number of paragraphs
    for _ in range(num_paragraphs):
        num_sentences = random.randint(1, 20)  # Generate a random number of sentences per paragraph
        content = " ".join("".join(random.choices(string.ascii_letters, k=random.randint(3, 20))) for _ in range(num_sentences))
        doc.add_paragraph(content)

    doc.save(file_name)

# Generate 100 Word documents with random content
for i in range(100):
    file_name = f"output_{i+1}.docx"
    generate_word_document(file_name)