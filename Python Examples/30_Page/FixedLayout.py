from spire.doc import *
from spire.doc.common import *


def WriteAllText(fname: str, text: List[str]):
    fp = open(fname, "w", encoding='utf-8')
    for s in text:
        fp.write(s)
    fp.close()


inputFile = "Data/Template_Docx_3.docx"
outputFile = "FixedLayout.txt"

# Create a new instance of Document
doc = Document()

# Load the document from the specified file
doc.LoadFromFile(inputFile, FileFormat.Docx)

# Create a FixedLayoutDocument from the loaded document
layoutDoc = FixedLayoutDocument(doc)

# Get the first line in the first column of the first page
line = layoutDoc.Pages[0].Columns[0].Lines[0]

# Create a StringBuilder to store the output text
stringBuilder = []
stringBuilder.append("Line: " + line.Text + "\n")

# Get the paragraph that contains the line and append its text to the StringBuilder
para = line.Paragraph
stringBuilder.append("Paragraph text: " + para.Text + "\n")

# Get the text content of the first page
pageText = layoutDoc.Pages[0].Text
stringBuilder.append(pageText + "\n")

# Iterate through each page in the FixedLayoutDocument
for i in range(layoutDoc.Pages.Count):
    page = layoutDoc.Pages[i]
    # Get all the lines on the current page
    lines = page.GetChildEntities(LayoutElementType.Line, True)

    # Append the page index and number of lines to the StringBuilder
    stringBuilder.append("Page " + str(page.PageIndex) + " has " +
                         str(lines.Count) + " lines.\n")

# Append the lines of the first paragraph to the StringBuilder
# (except runs and nodes in the header and footer).
stringBuilder.append("The lines of the first paragraph:\n")
paragraphLines = layoutDoc.GetLayoutEntitiesOfNode(
    (doc.FirstChild).Body.Paragraphs[0])
for i in range(paragraphLines.Count):
    paragraphLine = paragraphLines.get_Item(i)
    stringBuilder.append(paragraphLine.Text.strip() + "\n")
    stringBuilder.append(str(paragraphLine.Rectangle) + "\n")

# Write the contents of the StringBuilder to a text file
WriteAllText(outputFile, "".join(stringBuilder))

# Dispose of the document object when finished using it
doc.Dispose()
