"""
-------------------------------------------------------------------------------------------------------------------------------------------
    Creates a Word Doc with citations from an Excel spreadsheet specified by the user
-------------------------------------------------------------------------------------------------------------------------------------------
"""

#import libraries
from docx import Document
from openpyxl import Workbook
from openpyxl import load_workbook
import datetime

"""
-------------------------------------------------------------------------------------------------------------------------------------------
    Row and Column numbers in Excel
-------------------------------------------------------------------------------------------------------------------------------------------
"""
ROW_DATA_STARTS = 2
#rows with data starts at 2

COL_NUM_AUTHORS = 1
COL_AUTH1_LAST_NAME = 2
COL_AUTH1_FIRST_NAME = 3
COL_AUTH1_MI = 4
COL_AUTH2_LAST_NAME = 5
COL_AUTH2_FIRST_NAME = 6
COL_AUTH2_MI = 7
COL_TITLE = 8
COL_CONTAINER = 9
COL_CONTRIBUTORS = 10
COL_VERSION = 11
COL_NUMBER = 12
COL_PUBLISHER = 13
COL_DATE_PUBLISHED = 14
COL_LOCATION = 15
COL_DATE_ACCESSED = 16
#to access data
#sheet.cell(row, COLUMN).value



"""
-------------------------------------------------------------------------------------------------------------------------------------------
    Class that holds the attributes for each citation and the functions for getting
    that data from spreadsheet
-------------------------------------------------------------------------------------------------------------------------------------------
"""
class Citation:
    def __init__(self):
        self.numAuthors = 0          #number of authors add et al. if more than 2
        self.authors = ""            #list of authors in string format
        self.title = ""              #title of article
        self.container = ""          #title of collection or website
        self.contributors = ""       #editors etc
        self.version = ""            #edition or version
        self.number = 0              #number or vol
        self.publisher = ""          #publisher
        self.location = ""           #page numbers or url
        self.datePublished = ""      #date published or updated 
        self.dateAccessed = ""       #date website accessed

    #Read from Excel and store date for publishing
    def inputDatePublished(self, row):
        date = sheet.cell(row, COL_DATE_PUBLISHED).value
        #if date is in datetime format convert into a string
        if(type(date) == datetime.datetime):
            self.datePublished = str(date.day) + " " + convertNumToMonth(date.month) + " " + str(date.year)
        else:
            self.datePublished = str(date)

    
    #Read from Excel and store accessed date
    def inputDateAccessed(self, row):
        date = sheet.cell(row, COL_DATE_ACCESSED).value
        #if date is in datetime format convert into a string
        if(type(date) == datetime.datetime):
            self.dateAccessed = "Acessed " + str(date.day) + " " + convertNumToMonth(date.month) + " " + str(date.year)
        else:
            self.dateAcessed = "Accessed " + str(date)
    

        

"""
-------------------------------------------------------------------------------------------------------------------------------------------
    Function to get the name of the output file from the user
-------------------------------------------------------------------------------------------------------------------------------------------
"""
def inputExcelFileName():
    
    fileExtension = ".xlsx"
    
    #get a name of file from user
    fileName = input("Enter name of Excel file: ")
    
    #append .xlsx if input name doesn't have it
    if (fileName.endswith(fileExtension) == False):
        fileName += fileExtension

    #return the name of the file
    return fileName


"""
-------------------------------------------------------------------------------------------------------------------------------------------
    Function to get the name of the output file from the user
-------------------------------------------------------------------------------------------------------------------------------------------
"""
def inputCitationFileName():
    
    #default name of the document
    fileName = "citations.docx"
    fileExtension = ".docx"

    #get a name of document from user
    fileName = input("Where to save citations: ")
    
    #append .docx if input name doesn't have it
    if (fileName.endswith(fileExtension) == False):
        fileName += fileExtension

    #return the name of the file
    return fileName


"""
-------------------------------------------------------------------------------------------------------------------------------------------
    Function to return a month string based on number
-------------------------------------------------------------------------------------------------------------------------------------------
"""
def convertNumToMonth(num):
    if num == 1:
        return "Jan"
    elif num == 2:
        return "Feb"
    elif num == 3:
        return "Mar"
    elif num == 4:
        return "Apr"
    elif num == 5:
        return "May"
    elif num == 6:
        return "Jun"
    elif num == 7:
        return "Jul"
    elif num == 8:
        return "Aug"
    elif num == 9:
        return "Sep"
    elif num == 10:
        return "Oct"
    elif num == 11:
        return "Nov"
    elif num == 12:
        return "Dec"
    else:
        return "???"
            


"""
-------------------------------------------------------------------------------------------------------------------------------------------
    Main function execution
-------------------------------------------------------------------------------------------------------------------------------------------
"""

print("MLA Citation Maker")    

excelFileOpened = True

excelFileName = inputExcelFileName()
#try to open the excel sheet
try:
    #load workbook need to import library
    workbook = load_workbook(filename = excelFileName)
    sheetName = workbook.sheetnames[0]
    sheet = workbook[sheetName]
except:
    excelFileOpened = False
    print("Failed to open Excel sheet ")
    #pause for user to see error message
    input("Press Enter to quit")



#Continue if the excel file opened
if excelFileOpened:
    
    documentName = inputCitationFileName()
    #try to open the document or create it if it doesn't exist
    try:
        document = Document(documentName)
    except:
        document = Document()

    print("Succesfully opened both")
    document.add_paragraph("Works Cited")


    citation = Citation()
    citation.inputDatePublished(4)
    citation.inputDateAccessed(4)

    print(citation.datePublished)
    print(citation.dateAccessed)
    
    
    document.save(documentName)











"""
#write heading
document.add_paragraph("Works Cited")

firstCitation = Citation()

firstCitation.inputAuthors()
p=document.add_paragraph(firstCitation.numAuthors)
p.add_run(firstCitation.authors[0])

#Save the document
document.save(documentName)
"""

"""
#read sheet
#load workbook need to import library
wb2 = load_workbook(filename = 'readSheet.xlsx')
#get the name of the first sheet
sheetName2 = wb2.sheetnames[0]
#connect the worksheet object using the sheet name
worksheet2 = wb2[sheetName2]
print(worksheet2.cell(row = 1, column = 1).value)
print(worksheet2.cell(row = 2, column = 2).value)
if (worksheet2.cell(row = 2, column = 2).value) == None:
    print("That cell was empty")
"""




