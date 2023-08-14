from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_2.docx"
outputFile =  "InsertPageBreakFirstApproach.docx"
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Find the specified word "technology" where we want to insert the page break.
selections = document.FindAllString("technology", True, True)
#Traverse each word "technology".
for ts in selections:
    range = ts.GetAsOneRange()
    paragraph = range.OwnerParagraph
    index = paragraph.ChildObjects.IndexOf(range)
    #Create a new instance of page break and insert a page break after the word "technology".
    pageBreak = Break(document, BreakType.PageBreak)
    paragraph.ChildObjects.Insert(index + 1, pageBreak)
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
