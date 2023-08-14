from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/HeaderAndFooter.docx"
outputFile = "LockHeader.docx"

# Load the document
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first section
section = doc.Sections[0]

# Protect the document and set the ProtectionType as AllowOnlyFormFields
doc.Protect(ProtectionType.AllowOnlyFormFields, "123")

# Set the ProtectForm as false to unprotect the section
section.ProtectForm = False

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
