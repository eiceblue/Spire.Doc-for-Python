from spire.doc import *
from spire.doc.common import *

outputFile = "AddVariables.docx"

#Create Word document.
document = Document()
#Add a section.
section = document.AddSection()
#Add a paragraph.
paragraph = section.AddParagraph()
#Add a DocVariable field.
paragraph.AppendField("A1", FieldType.FieldDocVariable)
#Add a document variable to the DocVariable field.
document.Variables.Add("A1", "12")
#Update fields.
document.IsUpdateFields = True
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
