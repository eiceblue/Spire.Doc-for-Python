from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template_Docx_4.docx"
outputFile = "AddPageNumbersInSections.docx"

# Create Word document.
document = Document()

# Load the file from disk.
document.LoadFromFile(inputFile)

# Repeat step2 and Step3 for the rest sections, so change the code with for loop.
for i in range(0, 3):
    footer = document.Sections[i].HeadersFooters.Footer
    footerParagraph = footer.AddParagraph()
    footerParagraph.AppendField("page number", FieldType.FieldPage)
    footerParagraph.AppendText(" of ")
    footerParagraph.AppendField(
        "number of pages", FieldType.FieldSectionPages)
    footerParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

    if i == 2:
         break
    else:
        document.Sections[i + 1].PageSetup.RestartPageNumbering = True
        document.Sections[i + 1].PageSetup.PageStartingNumber = 1

# Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()

