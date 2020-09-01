import PyPDF2
import os
import time
import shutil
import sys


# Helper class used to map pages numbers to bookmarks


class BookmarkToPageMap(PyPDF2.PdfFileReader):

    def getDestinationPageNumbers(self):
        def _setup_outline_page_ids(outline, _result=None):
            if _result is None:
                _result = {}
            for obj in outline:
                if isinstance(obj, PyPDF2.pdf.Destination):
                    _result[(id(obj), obj.title)] = obj.page.idnum
                elif isinstance(obj, list):
                    _setup_outline_page_ids(obj, _result)
            return _result

        def _setup_page_id_to_num(pages=None, _result=None, _num_pages=None):
            if _result is None:
                _result = {}
            if pages is None:
                _num_pages = []
                pages = self.trailer["/Root"].getObject()["/Pages"].getObject()
            t = pages["/Type"]
            if t == "/Pages":
                for page in pages["/Kids"]:
                    _result[page.idnum] = len(_num_pages)
                    _setup_page_id_to_num(page.getObject(), _result, _num_pages)
            elif t == "/Page":
                _num_pages.append(1)
            return _result

        outline_page_ids = _setup_outline_page_ids(self.getOutlines())
        page_id_to_page_numbers = _setup_page_id_to_num()

        result = {}
        for (_, title), page_idnum in outline_page_ids.items():
            result[title] = page_id_to_page_numbers.get(page_idnum, '???')
        return result


def main(arg1, arg2, arg3, arg4):
    ################
    # Main Program #
    ################
    # Set parameters
    sourcePDFFile = arg1
    outputPDFDir = arg2
    outputNamePrefix = arg3
    deleteSourcePDF = arg4
    targetPDFFile = 'temppdfsplitfile.pdf'  # Temporary file

# I'm commenting it out, because we're running on Linux.
#    if outputPDFDir:
#        # Append backslash to output dir if necessary
#        if not outputPDFDir.endswith('\\'):
#            outputPDFDir = outputPDFDir + '\\'

    print('Parameters:')
    print(sourcePDFFile)
    print(outputPDFDir)
    print(outputNamePrefix)
    print(targetPDFFile)

    if os.path.exists(sourcePDFFile):
        print('Found source PDF file')
        # Copy file to local working directory
        shutil.copy(sourcePDFFile, targetPDFFile)

        # Process file
        pdfFileObj2 = open(targetPDFFile, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj2)
        pdfFileObj = BookmarkToPageMap(pdfFileObj2)

        # Get total pages
        numberOfPages = pdfReader.numPages
        print('PDF # Pages: ' + str(numberOfPages))

        i = 0
        newPageNum = 0
        prevPageNum = 0
        newPageName = ''
        prevPageName = ''
        for p, t in sorted([(v, k) for k, v in pdfFileObj.getDestinationPageNumbers().items()]):
            template = '%-5s  %s'
            print(template % ('Page', 'Title'))
            print(template % (p+1, t))

            newPageNum = p + 1
            newPageName = t

            if prevPageNum == 0 and prevPageName == '':
                print('First Page...')
                prevPageNum = newPageNum
                prevPageName = newPageName
            else:
                if newPageName:
                    print('Next Page...')
                    pdfWriter = PyPDF2.PdfFileWriter()
                    page_idx = 0
                    for i in range(prevPageNum, newPageNum):
                        pdfPage = pdfReader.getPage(i-1)
                        pdfWriter.insertPage(pdfPage, page_idx)
                        print('Added page to PDF file: ' + prevPageName + ' - Page #: ' + str(i))
                        page_idx += 1

#                   pdfFileName = outputNamePrefix + str(str(prevPageName).replace(':','_')).replace('*','_') + '.pdf'
                    # I've added "/" as a character to replace, because, again, we're running on Linux
                    pdfFileName = outputNamePrefix + \
                        str(str(prevPageName).replace(':', '_')).replace(
                            '*', '_').replace('/', ' ') + '.pdf'
                    pdfOutputFile = open(os.path.join(outputPDFDir, pdfFileName), 'wb')
                    pdfWriter.write(pdfOutputFile)
                    pdfOutputFile.close()
                    print('Created PDF file: ' + outputPDFDir + pdfFileName)

            i = prevPageNum
            prevPageNum = newPageNum
            prevPageName = newPageName

        # Split the last page
        print('Last Page...')
        pdfWriter = PyPDF2.PdfFileWriter()
        page_idx = 0
        for i in range(prevPageNum, numberOfPages + 1):
            pdfPage = pdfReader.getPage(i-1)
            pdfWriter.insertPage(pdfPage, page_idx)
            print('Added page to PDF file: ' + prevPageName + ' - Page #: ' + str(i))
            page_idx += 1

#       pdfFileName = outputNamePrefix + str(str(prevPageName).replace(':','_')).replace('*','_') + '.pdf'
        # I've added "/" as a character to replace, because, again, we're running on Linux
        pdfFileName = outputNamePrefix + \
            str(str(prevPageName).replace(':', '_')).replace('*', '_').replace('/', ' ') + '.pdf'
        pdfOutputFile = open(os.path.join(outputPDFDir, pdfFileName), 'wb')
        pdfWriter.write(pdfOutputFile)
        pdfOutputFile.close()
        print('Created PDF file: ' + outputPDFDir + pdfFileName)

        pdfFileObj2.close()

        # Delete temp file
        os.unlink(targetPDFFile)

        if newPageName:
            if deleteSourcePDF == True or deleteSourcePDF == "True":
                os.unlink(sourcePDFFile)


mypath = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Bundled CVs')
destination = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CVs')


onlyfiles = [os.path.join(mypath, f) for f in os.listdir(
    mypath) if os.path.isfile(os.path.join(mypath, f))]

for x in onlyfiles:
    main(x, destination, '', False)
