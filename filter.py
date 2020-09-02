import PyPDF2
import re
import string
import pandas as pd
import os
import glob
from tqdm import tqdm

# Defining the path & selecting files
cvs = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CVs')

pdfFiles = []
for filename in os.listdir(cvs):
    if filename.lower().endswith('.pdf'):
        pdfFiles.append(os.path.join(cvs, filename))

# Create dictionary with keywords
terms = {'Roche Experience': ['roche', 'cobas'],
         'Technical Experience': ['ivd', 'in vitro', 'dia', 'diagnostics'],
         'Education': ['biology', 'chemistry', 'english', 'arabic'],
         'Location': ['abha', 'khamis', 'khamis mushait', 'tendaha', 'ahad rafidah', 'fayfa', 'sabya',
                      'abu arish', 'jazan', 'najran', 'samtah', 'yadamah', 'al namas', 'bisha', 'tathleeth',
                      'al qundudhah', 'al bahah'],
         'Cometitors': ['mediserv', 'philips healthcare', 'beckman coulter', 'siemens healthineers',
                        'becton dickinson ', 'medtronic', '3m health care', 'biomérieux', 'ge healthcare',
                        'stago', 'mindray', 'sysmex', 'ortho', 'agfa healthcare', 'thermofisher', 'danaher',
                        'snibe', 'diasorin', 'biorad', 'cepheid', 'agilent technologies', 'sakura-finetek',
                        'biocare', 'biogenex', 'radiometer', 'instrumentation laboratory (il)', 'nova',
                        'mitsubishi', '77 electronika', ' j&j ', 'olympus', 'al faisaliah medical systems',
                        'amico', 'banaja medical co.', 'gulf medical co. ltd.',
                        'abdulla fouad group - medical supplies division', 'abdulrehman algosaibi gtc.',
                        'abdl rauof ibrahim batterjee & brothers company', 'al jeel medical', 'salehiya medical',
                        'samir photographic supplies ltd.', 'arabian health care supply company, olayan group',
                        'abbott laboratories', ' bd '],
         'Irrelevant Candidates': ['pharmaceutical', 'pharmacy', 'transferable', 'visa', 'iqama']}

# prepare a summary table
candidates = {'Name': [], 'Text': []}
for k in list(terms):
    candidates[k] = []

########################################################################
# Extract text from every page on the file

for x in tqdm(pdfFiles):
    pdfFileObj = open(x, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    num_pages = pdfReader.numPages

    count = 0
    text = ""

    while count < num_pages:
        pageObj = pdfReader.getPage(count)
        count += 1
        text += pageObj.extractText()

    # Convert all strings to lowercase
    text = text.lower()

    # Remove numbers
    text = re.sub(r'\d+', '', text)

    # Remove punctuation
    text = text.translate(str.maketrans('', '', string.punctuation))
#########################################################################

    y = 0
    # Obtain the scores for each area
    for k, v in terms.items():
        for word in v:
            if word in text:
                y += 1
        candidates[k].append(y)
        y = 0
    candidates['Name'].append('=HYPERLINK("{}", "{}")'.format(
        os.path.join('CVs', os.path.basename(x)), os.path.basename(x)))
    candidates['Text'].append(text)


# Create a dataframe for the summary table
summary = pd.DataFrame(candidates)
summary.set_index('Name', inplace=True)

# Export summary table
# summary.to_csv('screening results.csv')
try:
    summary.to_csv('results.csv')
except(FileCreateError):
    print('de')

    # xlsxwriter.exceptions.FileCreateError
    # Try Catch Final
