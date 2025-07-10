from spire.doc import *
from spire.doc.common import *

inputFile = "Data/Template_Docx_1.docx"
outputFile = "SetFirstLineIndentChars.docx"

# Create a new Document object
document = Document()

# Load a Word document from the specified file path
document.LoadFromFile(inputFile)

# Create a Paragraph object using the loaded document
para = Paragraph(document)

# Append text to the paragraph and customize its formatting
textRange1 = para.AppendText("This is an inserted paragraph.")
textRange1.CharacterFormat.TextColor = Color.get_Blue()
textRange1.CharacterFormat.FontSize = 15

# Set the first line indent to 0 characters
para.Format.SetFirstLineIndentChars(0)

# Insert the paragraph at index 1 in the first section of the document
document.Sections[0].Paragraphs.Insert(1, para)

document.SaveToFile(outputFile, FileFormat.Docx2013)

# Dispose the Document object to release resources
document.Dispose()
