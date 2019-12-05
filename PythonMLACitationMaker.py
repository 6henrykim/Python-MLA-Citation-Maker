'''
    Creates MLA citation page based on input from user
'''

#import the docx library
from docx import Document

#open the document
document = Document()

#write heading
document.add_paragraph("Citations")

#Save the document
document.save("test.docx")
