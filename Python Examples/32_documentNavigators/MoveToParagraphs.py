from spire.doc import *

# Create a new empty document instance.
doc = Document()

# Create a document navigator to help navigate and modify the document content.
navigator = DocumentNavigator(doc)

# Load an existing Word document from the specified relative file path.
doc.LoadFromFile("Data\\Sample.docx")

# Move the navigator's cursor to the first section of the document (section index 0).
navigator.MoveToSection(0)

# Move the cursor to the third paragraph (index 2) within the current section, at character offset 0.
navigator.MoveToParagraph(2, 0)

# Insert new text at the current cursor position, overwriting any existing content from that point onward.
navigator.Write("This is new content......")

# Save the modified document to a new file.
doc.SaveToFile("MoveToParagraphs.docx", FileFormat.Docx)

# Close the document to release internal resources.
doc.Close()

# Explicitly dispose of the document object to free memory immediately.
doc.Dispose()