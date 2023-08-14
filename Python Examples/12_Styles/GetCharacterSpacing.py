from io import FileIO
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Insert.docx"
outputFile = "GetCharacterSpacing.txt"
       

#Create a document
document = Document()

#Load the document from disk.
document.LoadFromFile(inputFile)

#Get the first section of document
section = document.Sections[0]

#Get the first paragraph 
paragraph = section.Paragraphs[0]

#Define two variables
fontName = ""
fontSpacing = 0

#Traverse the ChildObjects 
for i in range(paragraph.ChildObjects.Count):
    docObj = paragraph.ChildObjects.get_Item(i)
    #If it is TextRange
    if isinstance(docObj, TextRange):
        textRange = docObj if isinstance(docObj, TextRange) else None

        #Get the font name
        fontName = textRange.CharacterFormat.FontName

        #Get the character spacing
        fontSpacing = textRange.CharacterFormat.CharacterSpacing
content = "The font of first paragraph is " + fontName + ", the character spacing is " + str(fontSpacing) + "pt."
#Write the contents in a TXT file
with FileIO(outputFile, mode="w") as f:
    f.write(content.encode("utf-8"))
