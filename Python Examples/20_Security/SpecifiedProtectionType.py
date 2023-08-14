from spire.doc import *
from spire.doc.common import *

outputFile = "SpecifiedProtectionType.docx"
inputFile = "./Data/Template_Docx_2.docx"

# Create Word document.
document = Document()

# Load the file from disk.
document.LoadFromFile(inputFile)

# Protect the Word file.
document.Protect(ProtectionType.AllowOnlyReading, "123456")

# Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
