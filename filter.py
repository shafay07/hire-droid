import PyPDF2
import re
import string
import pandas as pd
import os
import glob
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
    cvs = Path("batch_cvs")
    # Indivisual batch files in a list
    pdfFiles = []
    for file in cvs.iterdir():
        pdfFiles.append(file)
    print('Done...\n')

    # Loading keywords.json
    print('***Loading keywords***')
    with open('keywords.json') as json_file:
        keywords = json.load(json_file)
    print('Done...\n')

    return pdfFiles, keywords


# get pdfBookmarks
def getPdfBookmarks(batch_file):
    pdf_file_obj = open(batch_file, 'rb')
    pdf_reader = PyPDF2.PdfFileReader(pdf_file_obj)
    bookmarks = pdf_reader.outlines
    return pdf_reader, bookmarks


# extracts the text for each resume
def getEachResumeText(current_batch, pdf_reader, bookmarks, bookmarks_len, start_bookmark=0):
    start_page = pdf_reader.getDestinationPageNumber(
        bookmarks[start_bookmark])
    # at last bookmark set end_page to last page of the document
    if bookmarks_len == start_bookmark:
        end_page = pdf_reader.getNumPages()
    # else set it to next bookmark
    else:
        end_bookmark = start_bookmark+1
        end_page = pdf_reader.getDestinationPageNumber(
            bookmarks[end_bookmark])

    # itterate over a complete resume to extract text
    clean_extracted = ''
    for page in range(start_page, end_page):
        with pdfplumber.open(current_batch) as pdf:
            current_page = pdf.pages[page]
            extracted = current_page.extract_text()
            cleaned = clean(extracted, lower=True, no_line_breaks=True, no_phone_numbers=True,
                            no_emails=True, no_urls=True, no_numbers=True, no_digits=True)
            clean_extracted += cleaned
    # helper function to save whats extracted by pdfplumber in a text file
    saveCleanText(clean_extracted)
    return clean_extracted, start_page


# Itterate over each resume for each batch to search the matching keywords
def searchEachResume(batch_pdf, keywords):
    # set csv header for results.csv
    setCsvHeader(keywords.keys())
    # for each batch
    for current_batch in batch_pdf:
        pdfReader, bookmarks = getPdfBookmarks(current_batch)
        bookmarks_len = getBookmarkLength(bookmarks)
        # for each resume in a batch
        for resume_index, resume in enumerate(tqdm(bookmarks, bar_format='{l_bar}{bar:50}{r_bar}{bar:-10b}')):
            # for resume_index, resume in enumerate(tqdm(range(96, 98), bar_format='{l_bar}{bar:50}{r_bar}{bar:-10b}')):
            clean_extracted, start_page = getEachResumeText(
                current_batch, pdfReader, bookmarks, bookmarks_len, resume_index)
            # get hits for each keyword
            # create a hit list
            hits = {}
            for key, value in keywords.items():
                for val in value:
                    count = clean_extracted.count(val)
                    hits["Applicant id"] = '=HYPERLINK("{}#page={}","{}")'.format(
                        current_batch, start_page+1, resume_index+1)
                    hits[key] = count
            saveCsvResults(hits)


# Helper function to get total number of pages
def getBookmarkLength(bookmarks):
    return len(bookmarks)-1


# Helper function to set header for csv file
def setCsvHeader(header):
    list_header = list(header)
    list_header.insert(0, 'Applicant ID')
    df = pd.DataFrame(columns=list_header)
    df.to_csv('results.csv', index=False)


# Helper function to write a text file to analyze what is extracted
def saveCleanText(to_write):
    file1 = open("clean_extracted.txt", "a",
                 encoding='utf-8')  # append mode
    file1.write("\n"+to_write+"\n")
    file1.write("***NEXT RESUME***")
    file1.close()


# Save final results of keywords matched in a csv
def saveCsvResults(to_append):
    df = pd.DataFrame(to_append, index=[0])
    df.to_csv('results.csv', mode='a', header=False, index=False)


if __name__ == "__main__":
    banner = pyfiglet.figlet_format("Hire - Droid", font="standard")
    print(banner)
    batchPdf, keywords = loadFiles()
    searchEachResume(batchPdf, keywords)
    print('Done...')
    print('\nPlease find results.csv file, thank you for using Hire-Droid.')
