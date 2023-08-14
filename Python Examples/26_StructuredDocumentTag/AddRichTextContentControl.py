from spire.doc import *
from spire.doc.common import *

outputFile = "AddRichTextContentControl.docx"

#Create a document
document = Document()

#Add a new section.
section = document.AddSection()

#Add a paragraph
paragraph = section.AddParagraph()

#Append textRange for the paragraph
txtRange = paragraph.AppendText("The following example shows how to add RichText content control in a Word document. \n")

#Append textRange 
txtRange = paragraph.AppendText("Add RichText Content Control:  ")

#Set the font format
txtRange.CharacterFormat.Italic = True

#Create StructureDocumentTagInline for document
sdt = StructureDocumentTagInline(document)

#Add sdt in paragraph
paragraph.ChildObjects.Add(sdt)

#Specify the type
sdt.SDTProperties.SDTType = SdtType.RichText

#Set displaying text
text = SdtText(True)
text.IsMultiline = True
sdt.SDTProperties.ControlProperties = text

#Crate a TextRange
rt = TextRange(document)
rt.Text = "Welcome to use "
rt.CharacterFormat.TextColor = Color.get_Green()
sdt.SDTContent.ChildObjects.Add(rt)
rt = TextRange(document)
rt.Text = "Spire.Doc"
rt.CharacterFormat.TextColor = Color.get_OrangeRed()
sdt.SDTContent.ChildObjects.Add(rt)

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
