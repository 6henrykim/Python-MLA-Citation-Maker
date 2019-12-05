"""
    Creates MLA citation page based on input from user
"""

#import the docx library
from docx import Document


class Citation:
    numAuthors = 0
    authors = []
    title = ""
    container = ""
    contributors = []
    version = ""
    publisher = ""
    location = ""
    datePublished = ""
    dateAccessed = ""

    def inputAuthors(self):
        self.numAuthors = input("How many authors? ")

        
        lastName = input("Enter last name: ")
        firstName = input("Enter first name: ")
        middleInitial = input("Enter middle initial: ")
        
        self.authors.append(lastName + ", " + firstName + " " + middleInitial + ".")
    
    

#try to open the document or create it if it doesn't exist
try:
    document = Document("test.docx")
except:
    document = Document()

firstCitation = Citation()

#write heading
document.add_paragraph("Works Cited")

firstCitation.inputAuthors()
p=document.add_paragraph(firstCitation.numAuthors)
p.add_run(firstCitation.authors[0])

#Save the document
document.save("test.docx")
