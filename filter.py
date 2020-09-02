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
def getEachResumeText(current_batch, pdf_reader, bookmarks, start_bookmark=0):
    start_page = pdf_reader.getDestinationPageNumber(
        bookmarks[start_bookmark])
    end_bookmark = start_bookmark+1
    end_page = pdf_reader.getDestinationPageNumber(
        bookmarks[end_bookmark])
    print("startPage {}, endPage {}".format(start_page, end_page))

    # itterate over a complete resume to extract text
    for page in range(start_page, end_page):
        clean_extracted = ''
        with pdfplumber.open(current_batch) as pdf:
            current_page = pdf.pages[page]
            extracted = current_page.extract_text()
            cleaned = clean(extracted, lower=True, no_line_breaks=True, no_phone_numbers=True,
                            no_emails=True, no_urls=True, no_numbers=True, no_digits=True)
            clean_extracted += cleaned
        # Debug code to see whats extracted by pdfplumber
        '''
        file1 = open("clean_extracted.txt", "a",
                     encoding='utf-8')  # append mode
        file1.write(clean_extracted)
        file1.write("\n")
        file1.write("***NEXT RESUME***")
        file1.close()
        '''
    return clean_extracted
    #     while count < num_pages:
    #         pageObj = pdfReader.getPage(count)
    #         count += 1
    #         text += pageObj.extractText()

    #     # Convert all strings to lowercase
    #     text = text.lower()

    #     # Remove numbers
    #     text = re.sub(r'\d+', '', text)

    #     # Remove punctuation
    #     text = text.translate(str.maketrans('', '', string.punctuation))
    # #########################################################################

    #     y = 0
    #     # Obtain the scores for each area
    #     for k, v in terms.items():
    #         for word in v:
    #             if word in text:
    #                 y += 1
    #         candidates[k].append(y)
    #         y = 0
    #     candidates['Name'].append('=HYPERLINK("{}", "{}")'.format(
    #         os.path.join('extracted_cvs', os.path.basename(x)), os.path.basename(x)))
    #     candidates['Text'].append(text)

    # print('Done...')
    # print('Generating output file...')
    # # Create a dataframe for the summary table
    # summary = pd.DataFrame(candidates)
    # summary.set_index('Name', inplace=True)
    # # Export summary table
    # summary.to_csv('results.csv')
    print('Done...')
    print('Please find results.csv file, thank you for using Hire-Droid.')


if __name__ == "__main__":
    banner = pyfiglet.figlet_format("Hire - Droid", font="standard")
    print(banner)
    batchPdf, keywords = loadFiles()
    for batch_index, current_batch in enumerate(batchPdf):
        pdfReader, bookmarks = getPdfBookmarks(current_batch)
        for resume_index, resume in enumerate(tqdm(bookmarks, bar_format='{l_bar}{bar:50}{r_bar}{bar:-10b}')):
            if resume_index == 20:
                exit()
            print("\nProcessing batch: {}, Resume: {}".format(
                batch_index+1, resume_index+1))
            clean_extracted = getEachResumeText(
                current_batch, pdfReader, bookmarks, resume_index)
