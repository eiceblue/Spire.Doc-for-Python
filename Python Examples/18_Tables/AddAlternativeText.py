from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/TableSample.docx"
outputFile = "AddAlternativeText.docx"

# Load the document
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first section
section = doc.Sections[0]

# Get the first table in the section
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Add alternative text
# Add title
table.Title = "Table 1"
# Add description
table.TableDescription = "Description Text"

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()

