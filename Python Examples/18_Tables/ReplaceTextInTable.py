from spire.doc import *
from spire.doc.common import *

outputFile = "ReplaceTextInTable.docx"
inputFile = "./Data/ReplaceTextInTable.docx"

# Load Word from disk
doc = Document()
doc.LoadFromFile(inputFile)

# Get the first section
section = doc.Sections[0]

# Get the first table in the section
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Define a regular expression to match the {} with its content
regex = Regex("""{[^\\}]+\\}""")

# Replace the text of table with regex
table.Replace(regex, "E-iceblue")

# Replace old text with new text in table
table.Replace("Beijing", "Component", False, True)

# Save the Word document
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()
