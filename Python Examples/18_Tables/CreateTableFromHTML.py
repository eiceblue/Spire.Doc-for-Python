from spire.doc import *
from spire.doc.common import *


outputFile = "CreateTableFromHTML.docx"

# HTML string
HTML = "<table border='2px'>" + "<tr>" + "<td>Row 1, Cell 1</td>" + "<td>Row 1, Cell 2</td>" + \
    "</tr>" + "<tr>" + "<td>Row 2, Cell 2</td>" + \
    "<td>Row 2, Cell 2</td>" + "</tr>" + "</table>"

# Create a Word document
document = Document()

# Add a section
section = document.AddSection()

# Add a paragraph and append html string
section.AddParagraph().AppendHTML(HTML)

# Save to Word document
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
