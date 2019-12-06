"""
-------------------------------------------------------------------------------------------------------------------------------------------
    Creates a Word Doc with citations from an Excel spreadsheet specified by the user
-------------------------------------------------------------------------------------------------------------------------------------------
"""

#import libraries
from docx import Document
from docx.shared import Pt
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Length
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
#to access data:
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

#flags for whether the files opened properly
excelFileOpened = True
citationFileOpened = True

excelFileName = inputExcelFileName()
#try to open the Excel file
try:
    #load workbook 
    workbook = load_workbook(filename = excelFileName)
    #get the name of the first worksheet
    sheetName = workbook.sheetnames[0]
    #set the sheet using the name of the first worksheet
    sheet = workbook[sheetName]
except:
    excelFileOpened = False
    print("Failed to open \"" + excelFileName + "\"")
    #pause for user to see error message
    input("Press Enter to quit")


#Try to open the Word doc if the Excel file opened
if excelFileOpened:
    
    documentName = inputCitationFileName()
    #try to open the specified excel file
    try:
        document = Document(documentName)
        document.save(documentName)
    #display error message if file couldn't be opened and saved
    except PermissionError:
        citationFileOpened = False
        print("Failed to open \"" + documentName + "\": make sure no program has the file open")
        #pause for user to see error message
        input("Press Enter to quit")
    #otherwise create the file
    except:
        document = Document()
        document.save(documentName)

#Continue if both files opened
if excelFileOpened and citationFileOpened:
    print("Succesfully accessed both files")

    #Write the heading to the citation doc
    heading = document.add_paragraph()
    headingRun = heading.add_run("Works Cited")
    headingFont = headingRun.font
    headingFont.name = "Times New Roman"
    headingFont.size = Pt(12)
    headerFormatting = heading.paragraph_format
    headerFormatting.line_spacing_rule = 2   #set double spaceing
    headerFormatting.alignment = WD_ALIGN_PARAGRAPH.CENTER #center the heading
    headerFormatting.page_break_before = True    #put header on a new page

    #Citation Formatting
    paragraph = document.add_paragraph()
    run = paragraph.add_run("Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text Lots of Text ")
    font = run.font
    font.name = "Times New Roman"
    font.size = Pt(12)
    paragraphFormatting = paragraph.paragraph_format
    paragraphFormatting.line_spacing_rule = 2   #set double spaceing
    paragraphFormatting.keep_together = True    #keeps citation on same page
    paragraphFormatting.left_indent = Inches(0.5)   #indent citations
    paragraphFormatting.first_line_indent = Inches(-0.5)    #negative to make fist line hanging indent
    









    #Save the file
    document.save(documentName)




"""
citation = Citation()
citation.inputDatePublished(4)
citation.inputDateAccessed(4)

print(citation.datePublished)
print(citation.dateAccessed)
"""


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




