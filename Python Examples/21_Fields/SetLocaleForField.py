from spire.doc import *
from spire.doc.common import *

outputFile = "SetLocaleForField.docx"
inputFile = "./Data/SampleB_2.docx"

#Open a Word document
document = Document()
document.LoadFromFile(inputFile)

#Get the first section
section = document.Sections[0]

par = section.AddParagraph()

#Add a date field
field = par.AppendField("DocDate", FieldType.FieldDate)

#Set the LocaleId for the textrange
( field.OwnerParagraph.ChildObjects[0] if isinstance(field.OwnerParagraph.ChildObjects[0], TextRange) else None).CharacterFormat.LocaleIdASCII = 1049

field.FieldText = "2019-10-10"
#Update field
document.IsUpdateFields = True

document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

