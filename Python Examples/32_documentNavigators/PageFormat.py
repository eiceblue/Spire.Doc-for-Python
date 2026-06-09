from spire.doc import *

# Create a new empty document instance.
doc = Document()

# Create a document navigator to help navigate and modify the document content.
navigator = DocumentNavigator(doc)

# Load an existing Word document from the specified relative file path.
doc.LoadFromFile("Data\\Sample.docx")

# Move the navigator's cursor to the first section (section index 0) of the document.
navigator.MoveToSection(0)

# Set the page margins for the current section.
navigator.PageSetup.Margins = MarginsF(100.0, 80.0, 100.0, 80.0)

# Set the page size of the current section to Letter.
navigator.PageSetup.PageSize = PageSize.Letter()

# Save the modified document to a new file.
doc.SaveToFile("PageFormat.docx", FileFormat.Docx)

# Close the document to release internal resources.
doc.Close()

# Explicitly dispose of the document object to free memory immediately.
doc.Dispose()