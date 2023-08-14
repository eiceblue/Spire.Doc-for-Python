from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/MultiplePages.docx"
outputFile = "DifferentFirstPage.docx"

# Load the document
doc = Document()
doc.LoadFromFile(inputFile)

# Get the section and set the property true
section = doc.Sections[0]
section.PageSetup.DifferentFirstPageHeaderFooter = True

# Set the first page header. Here we append a picture in the header
paragraph1 = section.HeadersFooters.FirstPageHeader.AddParagraph()
paragraph1.Format.HorizontalAlignment = HorizontalAlignment.Right

headerimage = paragraph1.AppendPicture("./Data/E-iceblue.png")

# Set the first page footer
paragraph2 = section.HeadersFooters.FirstPageFooter.AddParagraph()
paragraph2.Format.HorizontalAlignment = HorizontalAlignment.Center
FF = paragraph2.AppendText("First Page Footer")
FF.CharacterFormat.FontSize = 10

# Set the other header & footer. If you only need the first page header & footer, don't set this
paragraph3 = section.HeadersFooters.Header.AddParagraph()
paragraph3.Format.HorizontalAlignment = HorizontalAlignment.Center
NH = paragraph3.AppendText("Spire.Doc for Python")
NH.CharacterFormat.FontSize = 10

paragraph4 = section.HeadersFooters.Footer.AddParagraph()
paragraph4.Format.HorizontalAlignment = HorizontalAlignment.Center
NF = paragraph4.AppendText("E-iceblue")
NF.CharacterFormat.FontSize = 10

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
