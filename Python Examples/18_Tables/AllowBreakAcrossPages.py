from spire.doc import *
from spire.doc.common import *


# Create word document
inputFile = "./Data/AllowBreakAcrossPages.docx"
outputFile = "AllowBreakAcrossPages.docx"

document = Document()
document.LoadFromFile(inputFile)

section = document.Sections[0]
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

for i in range(table.Rows.Count):
    row = table.Rows.get_Item(i)
    # Allow break across pages
    row.RowFormat.IsBreakAcrossPages = True

# Save the Word document
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
