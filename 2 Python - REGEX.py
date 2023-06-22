import os
import re
import multiprocessing
from docx import Document
from glob import glob
import time

folder_path = os.path.join(os.path.expanduser("~"), "Desktop/PNI letters")
keywords = (
    r"(?i)(\bhip replacement\b.*\bsciatic nerve\b|\bsciatic nerve\b.*\bhip replacement\b|"
    r"\bTHR\b.*\bsciatic nerve\b|\bsciatic nerve\b.*\bTHR\b)"
)

def search_keywords_in_file(file_path):
    try:
        # Open the Word document
        doc = Document(file_path)

        # Extract the content
        paragraphs = [p.text for p in doc.paragraphs]
        content = " ".join(paragraphs)

        # Search for the keywords using regular expressions
        matches = re.findall(keywords, content)

        # Return the number of matches
        return len(matches)
    except Exception as e:
        print(f"Error processing file: {file_path}\n{e}")
        return 0

def process_file(file_path):
    result = search_keywords_in_file(file_path)
    filename = os.path.basename(file_path)
    return (filename, result)

if __name__ == "__main__":
    # Get a list of Word document file paths in the folder
    file_paths = glob(os.path.join(folder_path, "*.docx"))

    # Start recording script execution time
    start_time = time.time()

    # Process the files concurrently using multiple processes
    with multiprocessing.Pool() as pool:
        results = pool.map(process_file, file_paths)

    # Save the results to a tab-delimited text file
    output_file = "keyword_matches.txt"
    with open(output_file, "w", encoding="utf-8") as file:
        total_matches = sum(result for _, result in results)
        file.write("Total Matches: {}\n".format(total_matches))
        file.write("Filename\tMatching Keywords\n")
        for filename, result in results:
            file.write(f"{filename}\t{result}\n")

    # Calculate and print the script execution time
    end_time = time.time()
    execution_time = end_time - start_time
    print("Script execution time:", execution_time, "seconds")
