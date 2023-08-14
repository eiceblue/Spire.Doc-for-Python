from spire.doc import *
from spire.doc.common import *

outputFile = "AddCheckBoxContentControl.docx"

#Create a document
document = Document()

#Add a new section.
section = document.AddSection()

#Add a paragraph
paragraph = section.AddParagraph()

#Append textRange for the paragraph
txtRange = paragraph.AppendText("The following example shows how to add CheckBox content control in a Word document. \n")

#Append textRange 
txtRange = paragraph.AppendText("Add CheckBox Content Control:  ")

#Set the font format
txtRange.CharacterFormat.Italic = True

#Create StructureDocumentTagInline for document
sdt = StructureDocumentTagInline(document)

#Add sdt in paragraph
paragraph.ChildObjects.Add(sdt)

#Specify the type
sdt.SDTProperties.SDTType = SdtType.CheckBox

#Set properties for control
scb = SdtCheckBox()
sdt.SDTProperties.ControlProperties = scb

#Add textRange format
tr = TextRange(document)
tr.CharacterFormat.FontName = "MS Gothic"
tr.CharacterFormat.FontSize = 12

#Add textRange to StructureDocumentTagInline
sdt.ChildObjects.Add(tr)

#Set checkBox as checked
scb.Checked = True

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()