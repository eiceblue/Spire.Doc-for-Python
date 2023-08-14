
from spire.doc import *
from spire.doc.common import *

outputFile = "AlterLanguageDictionary.docx"

#Create Word document.
document = Document()
#Add new section and paragraph to the document.
sec = document.AddSection()
para = sec.AddParagraph()
#Add a textRange for the paragraph and append some Peru Spanish words.
txtRange = para.AppendText("corrige según diccionario en inglés")
txtRange.CharacterFormat.LocaleIdASCII = 10250
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()