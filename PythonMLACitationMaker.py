"""
    Creates MLA citation page based on input from the user
"""

#import libraries
from docx import Document

"""
    Class that holds the attributes for each citation and the functions for getting
    that data from the user
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

    def inputAuthors(self):
        self.numAuthors = input("How many authors? ")
        
        lastName = input("Enter last name: ")
        firstName = input("Enter first name: ")
        middleInitial = input("Enter middle initial: ")
        
        self.authors.append(lastName + ", " + firstName + " " + middleInitial + ".")

"""
    Function to get the name of the output file from the user
"""
def inputCitationFileName():
    
    #default name of the document
    fileName = "citations.docx"
    fileExtension = ".docx"

    #get a name of document from user
    fileName = input("Where to save citations: ")
    
    #append .docx if input name doesn't have it
    if (fileName.endswith(fileExtension) != True):
        fileName += ".docx"

    #return the name of the file
    return fileName




"""
-------------------------------------------------------------------------------------------------------------------------------------------
    Main function execution
"""
print("MLA Citation Maker")    

documentName = inputCitationFileName()

#try to open the document or create it if it doesn't exist
try:
    document = Document(documentName)
except:
    document = Document()

#write heading
document.add_paragraph("Works Cited")

firstCitation = Citation()

firstCitation.inputAuthors()
p=document.add_paragraph(firstCitation.numAuthors)
p.add_run(firstCitation.authors[0])

#Save the document
document.save(documentName)


