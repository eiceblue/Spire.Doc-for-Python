from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template.docx"
outputFile = "ImageHeaderAndFooter.docx"

# Load the document from disk
doc = Document()
doc.LoadFromFile(inputFile)

# Get the header of the first page
header = doc.Sections[0].HeadersFooters.Header

# Add a paragraph for the header
paragraph = header.AddParagraph()

# Set the format of the paragraph
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Right

# Append a picture in the paragraph

headerimage = paragraph.AppendPicture("./Data/E-iceblue.png")

headerimage.VerticalAlignment = ShapeVerticalAlignment.Bottom

# Get the footer of the first section
footer = doc.Sections[0].HeadersFooters.Footer

# Add a paragraph for the footer
paragraph2 = footer.AddParagraph()

# Set the format of the paragraph
paragraph2.Format.HorizontalAlignment = HorizontalAlignment.Left

# Append a picture in the paragraph

footerimage = paragraph2.AppendPicture("./Data/logo.png")

# Append text in the paragraph
TR = paragraph2.AppendText(
    "Copyright © 2013 e-iceblue. All Rights Reserved.")
TR.CharacterFormat.FontName = "Arial"
TR.CharacterFormat.FontSize = 10
TR.CharacterFormat.TextColor = Color.get_Black()

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
