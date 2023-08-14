from spire.doc import *
from spire.doc.common import *

outputFile = "Encrypt.docx"
inputFile = "./Data/Template.docx"

# Create word document
document = Document()

# Load Word document.
document.LoadFromFile(inputFile)

# encrypt document with password specified by textBox1
document.Encrypt("E-iceblue")

# Save as docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
