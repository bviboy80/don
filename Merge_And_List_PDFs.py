'''
Combine PDFs placed in a folder.
Combined PDF intended be printed duplex and sometimes stapled.

1. Count number of pages for each PDF to be merged. Add a blank
   to PDFs with odd page counts.
   
2. Combine the PDFs together.

3. Create a text file containing the following:
   - A list of each PDF and its page range within the combined PDF.
     (i.e., 1-20    Letter.pdf
            21-36   Notice.pdf)
   - The total number of pages in the combined PDF.
   - A comma separated list of the Page Ranges for Print Production
     to copy and paste into the OCE Page Programmer Software. 


Created by Shaun Thomas
Version 2.0
Created 7/26/2017

'''

import os
import sys
import PyPDF2


def main(pdfFolder):
    InputPath = os.path.abspath(pdfFolder)
    combinedPDF = os.path.join(InputPath, "combined.pdf")
    PDFandRangefile = os.path.join(InputPath, "PageList.txt")
    
    if os.path.exists(combinedPDF) or os.path.exists(PDFandRangefile):
        print "\nFiles already processed!"
    else:
        process(InputPath, combinedPDF, PDFandRangefile)

        
def process(InputPath, combinedPDF, PDFandRangefile):
    PDFfiles = os.listdir(InputPath)
    pageCountList = []    # Initialize list of page counts per PDF
    PDFandRangeList = []   # Initialize list of PDFs and page ranges
    merge = PyPDF2.PdfFileMerger()   # Initialize PDF page merger
    
    with open(PDFandRangefile, 'wb') as r:
        for file in PDFfiles:
            if os.path.isfile(file) and file[-4:].upper() == ".PDF":
                
                pdffile = os.path.join(InputPath, file)
                filename = os.path.basename(pdffile)
                
                document = PyPDF2.PdfFileReader(open(pdffile, 'rb'))
                blank = PyPDF2.PdfFileReader(open(r"P:\Donlin Recano\script\blank.pdf", 'rb'))
                
                docpagecount = addBlankPDF(merge, document, blank)
                PDFandRangeRow = calcPageRanges(filename, docpagecount, pageCountList)
            else:
                continue
            
            #Append info for each PDF to list
            PDFandRangeList.append(PDFandRangeRow)
        
        #Create combined PDF
        with open(combinedPDF, 'wb') as pdfout:
            merge.write(pdfout)
        
        createPDFandRangeFile(PDFandRangeList, r, pageCountList)

        
def addBlankPDF(merge, document, blank):
    '''Count pages of each PDF. If count is odd, add blank.
       Return page count.
    '''
    merge.append(fileobj=document)
    docpagecount =  document.getNumPages()
    if docpagecount % 2 > 0:
        merge.append(fileobj=blank)
        docpagecount+=1
    return docpagecount
    
def calcPageRanges(filename, docpagecount, pageCountList):
    '''Calculate page ranges for each PDF
       start = Sum the pages merged so far and add 1
       end = Add to the start: the page count of the current PDF minus 1 
    '''
    PDFandRangeRow = []
    rangestart = sum(pageCountList)+1
    rangeend = rangestart+docpagecount-1
    pageCountList.append(docpagecount)
    PDFandRangeRow.extend([filename, rangestart, rangeend])
    return PDFandRangeRow
    
def createPDFandRangeFile(PDFandRangeList, r, pageCountList):

    # Format the PDFandPageRange List
    for item in PDFandRangeList:
        print "\r\n{}-{}\t\t{}".format(item[1],item[2], item[0])
        r.write("{}-{}\t\t{}\r\n".format(item[1],item[2], item[0]))
    
    # Get the number of pages for the combined PDF.
    totals = sum(pageCountList)
    print "\r\n\r\nTOTAL PAGES:  {}".format(totals)
    r.write("\r\n\r\nTOTAL PAGES:  {}".format(totals))
    
    # Make a list of the page ranges for the Page Programmer.
    allpageranges = ""
    for item in PDFandRangeList:
        allpageranges += "{}-{}, ".format(item[1],item[2])
    r.write("\r\n\r\n\r\nRanges for Page Programmer:\r\n\r\n")
    r.write(allpageranges.rstrip(', '))


if __name__ == '__main__':
    main(sys.argv[1])
