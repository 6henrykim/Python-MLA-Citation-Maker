"""
    Creates MLA citation page Docx based on input from Excel Spreadsheet
"""

#import libraries
from docx import Document
from openpyxl import Workbook
from openpyxl import load_workbook


"""
-------------------------------------------------------------------------------------------------------------------------------------------
    Class that holds the attributes for each citation and the functions for getting
    that data from spreadsheet
-------------------------------------------------------------------------------------------------------------------------------------------
"""
class Citation:
    def __init__(self):
        self.numAuthors = 0          #add et al. if more than 2
        self.authors = []            #list for each author name stored as string
        self.title = ""              #title of article
        self.container = ""          #title of collection or website
        self.contributors = []       #editors etc
        self.version = ""            #edition or version
        self.number = 0              #number or vol
        self.publisher = ""          #publisher
        self.location = ""           #page numbers or url
        self.datePublished = ""      #date published or updated online
        self.dateAccessed = ""       #date website accessed

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




