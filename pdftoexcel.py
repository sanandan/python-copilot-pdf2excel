#create function to read a pdf file and convert it into excel format
import PyPDF2
import xlsxwriter
import os
import sys
import glob
def pdftoexcel():
    #read the pdf file
    pdfFileObj = open('AIR_Class-VI.pdf', 'rb')
    pdfReader = PyPDF2.PdfReader(pdfFileObj)
    #get the number of pages in the pdf file
    pages = len(pdfReader.pages)
    #create a list to store the data
    data = []
    for i in range(pages):
        #read the page
        pageObj = pdfReader.pages[i]
        #extract the text from the page
        text = pageObj.extract_text()
        #store the text in the list
        data.append(text)
    #close the pdf file
    pdfFileObj.close()
    #create a new excel file
    workbook = xlsxwriter.Workbook('test.xlsx')
    #create a new worksheet
    worksheet = workbook.add_worksheet()
    #set the row and column
    row = 0
    col = 0
    #write the data in the excel file
    for item in (data):
        worksheet.write(row, col, item)
        row += 1
    #close the excel file
    workbook.close()

if __name__ == '__main__':
    pdftoexcel()