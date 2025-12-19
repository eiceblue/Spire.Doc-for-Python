from spire.doc import *
from spire.doc.common import *

# Create a new Document object
document = Document()

# Load the document from the input file
document.LoadFromFile("Data/Insert.docx")

# When the system does not have the fonts used in a document installed, you can place the required fonts in a custom folder and then use setCustomFontsFolders to specify that the program should retrieve fonts from this path
document.SetCustomFontsFolders("D:\\Fonts")

# Save the document to the output file as PDF
document.SaveToFile("result.pdf", FileFormat.PDF)

# Close and dispose the document object
document.Close()
document.Dispose()