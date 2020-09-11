import string
import pandas as pd
import os
from tqdm import tqdm
from pathlib import Path
import pdfplumber
from cleantext import clean
import pyfiglet
import json


# Directory loader
def loadFiles():
    print('***Loading batch_cvs directory***')
    # Defining the path & selecting files
    cvs = Path("each_cvs")
    # Indivisual batch files in a list
    pdfFiles = []
    for file in cvs.iterdir():
        pdfFiles.append(file)
    print('Done...\n')

    # Loading keywords.json
    print('***Loading keywords***')
    with open(os.path.join(os.pardir, "keywords.json")) as json_file:
        keywords = json.load(json_file)
    print('Done...\n')

    return pdfFiles, keywords


# extracts the text for each resume
def getEachResumeText(resume):
    # itterate over a complete resume to extract text
    clean_extracted = ''
    with pdfplumber.open(resume) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            clean_extracted += clean(extracted, lower=True, no_line_breaks=True, no_phone_numbers=True,
                                     no_emails=True, no_urls=True, no_numbers=True, no_digits=True)
    # helper function to save whats extracted by pdfplumber in a text file
    saveCleanText(clean_extracted)
    return clean_extracted


# Itterate over each resume for each batch to search the matching keywords
def searchEachResume(listPdf, keywords):
    # set csv header for results.csv
    setCsvHeader(keywords.keys())
    # for each resume in a pdfList
    for resume_index, resume in enumerate(tqdm(listPdf, bar_format='{l_bar}{bar:50}{r_bar}{bar:-10b}')):
        clean_extracted = getEachResumeText(resume)
        # get hits for each keyword
        hits = {}
        hits["Applicant id"] = '=HYPERLINK("{}","{}")'.format(
            resume, resume_index+1)
        for key, value in keywords.items():
            count = 0
            for val in value:
                count += clean_extracted.count(val)
            hits[key] = count
        saveCsvResults(hits)


# Helper function to set header for csv file
def setCsvHeader(header):
    list_header = list(header)
    list_header.insert(0, 'Applicant ID')
    df = pd.DataFrame(columns=list_header)
    df.to_csv('each_results.csv', index=False)


# Helper function to write a text file to analyze what is extracted
def saveCleanText(to_write):
    file1 = open("each_clean_extracted.txt", "a",
                 encoding='utf-8')  # append mode
    file1.write("\n"+to_write+"\n")
    file1.write("***NEXT RESUME***")
    file1.close()


# Save final results of keywords matched in a csv
def saveCsvResults(to_append):
    df = pd.DataFrame(to_append, index=[0])
    df.to_csv('each_results.csv', mode='a', header=False, index=False)


if __name__ == "__main__":
    banner = pyfiglet.figlet_format("Hire - Droid", font="standard")
    print(banner)
    listPdf, keywords = loadFiles()
    searchEachResume(listPdf, keywords)
    print('Done...')
    print('\nPlease find each_results.csv file, thank you for using Hire-Droid.')
