from spire.doc import *
from spire.doc.common import *

document = Document()
#Load the document
document.LoadFromFile("./Data/Sample.docx")

#Find text
textSelections = document.FindAllString("Spire.Doc", False, True)

#test GetRanges()
textRanges = textSelections[0].GetRanges()
textRanges[0].CharacterFormat.HighlightColor = Color.get_Yellow()

#test GetAsRange()
textRange= textSelections[1].GetAsRange()
textRange[0].CharacterFormat.HighlightColor = Color.get_Red()

#test GetAsRange(bool IsCopyFormat)
textRange=textSelections[2].GetAsRange(True)
textRange[0].CharacterFormat.HighlightColor = Color.get_Green()

#Save the document.
document.SaveToFile("GetRanges.docx", FileFormat.Docx2016)
document.Close()