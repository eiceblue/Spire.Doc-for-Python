from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Insert.docx"
outputFile = "PreserveTheme.docx"

#Load the source document
doc = Document()
doc.LoadFromFile(inputFile)
#Create a new Word document
newWord = Document()
#Clone default style, theme, compatibility from the source file to the destination document
doc.CloneDefaultStyleTo(newWord)
doc.CloneThemesTo(newWord)
doc.CloneCompatibilityTo(newWord)
#Add the cloned section to destination document
newWord.Sections.Add(doc.Sections[0].Clone())
#Save the document
newWord.SaveToFile(outputFile, FileFormat.Docx)
newWord.Close()