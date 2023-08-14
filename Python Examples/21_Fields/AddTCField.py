from spire.doc import *
from spire.doc.common import *

outputFile = "AddTCField.docx"

# Create Word document.
document = Document()

# Add a new section.
section = document.AddSection()

# Add a new paragraph.
paragraph = section.AddParagraph()

# Add TC field in the paragraph
field = paragraph.AppendField("TC", FieldType.FieldTOCEntry)
field.Code = """TC """ + "\"Entry Text\"" + " \\f" + " t"
# Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
