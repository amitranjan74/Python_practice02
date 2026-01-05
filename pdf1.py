

import pdfplumber

# Path to your PDF file
pdf_file = r"D:\Desktop Backup\Python_Practice\GeminiPro_Project6\Chap12.pdf"

# Open the PDF
with pdfplumber.open(pdf_file) as pdf:
    all_text = ""
    for page in pdf.pages:
        # Extract text from each page
        text = page.extract_text()
        
        # Append the extracted text
        if text:
            all_text += text + "\n"

# Save the extracted text to a txt file
# Split all_text into a list of words
words = all_text.split()
print(len(words))
# Define the word range for each file
min_words = 3500
max_words = 3700

# Calculate how many chunks we need
total_words = len(words)
num_chunks = (total_words // min_words) + (1 if total_words % min_words != 0 else 0)

# Write the chunks to separate files
for i in range(num_chunks):
    start_index = i * min_words
    end_index = start_index + max_words
    chunk = words[start_index:end_index]
    
    # Create file names dynamically (e.g., extracted_text_part1.txt, part2.txt, etc.)
    file_name = f"extracted_text_part{i + 1}.txt"
    
    # Write the chunk to the file
    with open(file_name, "w", encoding="utf-8") as text_file:
        text_file.write(' '.join(chunk))
    
    print(f"Text extracted and saved to {file_name}")

