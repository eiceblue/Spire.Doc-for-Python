from spire.doc import *
from spire.doc.common import *


inputFile1 = "./Data/ReplaceContentWithDoc.docx"
inputFile2 = "./Data/Insert.docx"
outputFile = "ReplaceContentWithDoc.docx"

 #Create the first document
document1 = Document()

#Load the first document from disk.
document1.LoadFromFile(inputFile1)

#Create the second document
document2 = Document()

#Load the second document from disk.
document2.LoadFromFile(inputFile2)

#Get the first section of the first document 
section1 = document1.Sections[0]

#Create a regex
regex = Regex("""\\[MY_DOCUMENT\\]""", RegexOptions.none)

#Find the text by regex
textSections = document1.FindAllPattern(regex)

#Travel the found strings
for seletion in textSections:

    #Get the paragraph
    para = seletion.GetAsOneRange().OwnerParagraph

    #Get textRange
    textRange = seletion.GetAsOneRange()

    #Get the paragraph index
    index = section1.Body.ChildObjects.IndexOf(para)

    #Insert the paragraphs of document2
    for i in range(document2.Sections.Count):
        section2 = document2.Sections.get_Item(i)
        for j in range(section2.Paragraphs.Count):
            paragraph = section2.Paragraphs.get_Item(j)
            section1.Body.ChildObjects.Insert(index, paragraph.Clone() if isinstance(paragraph.Clone(), Paragraph) else None)
            #Remove the found textRange
            para.ChildObjects.Remove(textRange)

#Save the document.
document1.SaveToFile(outputFile, FileFormat.Docx)
document1.Dispose()

