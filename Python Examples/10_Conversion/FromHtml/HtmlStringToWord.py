
from spire.doc import *
from spire.doc.common import *

inputFile ="./Data/InputHtml.txt"
outputFile = "HtmlStringToWord.docx"


#Get html string.
with open(inputFile) as fp:
    HTML = fp.read()
#Create a new document.
document = Document()
#Add a section.
sec = document.AddSection()
#Add a paragraph and append html string.
sec.AddParagraph().AppendHTML(HTML)
#Save it to a Word file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()