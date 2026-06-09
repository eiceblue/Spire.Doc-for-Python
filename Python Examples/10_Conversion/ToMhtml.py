from spire.doc import *

# Create word document
document = Document()

# Load the file from disk.
document.LoadFromFile("data\\ToMhtml.docx")

# Save to RTF file.
document.SaveToFile("ToMhtml-out.mhtml", FileFormat.Mhtml)

# Close the document
document.Close()

# Dispose the document
document.Dispose()