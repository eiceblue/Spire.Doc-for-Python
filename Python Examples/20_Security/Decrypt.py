from spire.doc import *
from spire.doc.common import *

outputFile = "Decrypt.docx"
inputFile = "./Data/TemplateWithPassword.docx"

# Create word document
document = Document()
document.LoadFromFile(inputFile, FileFormat.Docx, "E-iceblue")

# Save as doc file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
