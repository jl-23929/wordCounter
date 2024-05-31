import re

# Sample text
text = " (2000)(2000) . (2000)."

# Define the regex pattern to match (Author, Year)

cleaned_text = re.sub(r'\.\s\(\d{4}\)\.', '', text)

print(cleaned_text)