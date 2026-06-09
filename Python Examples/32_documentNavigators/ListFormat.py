from spire.doc import *

# Create a new document instance.
doc = Document()

# Add a new section to the document.
section = doc.AddSection()

# Add the first paragraph to the section.
paragraph = section.AddParagraph()

# Append text to the first paragraph.
paragraph.AppendText("This is the first paragraph.")

# Add a second paragraph to the section.
paragraph = section.AddParagraph()

# Append text to the second paragraph.
paragraph.AppendText("This is the second paragraph.")

# Add a third paragraph to the section.
paragraph = section.AddParagraph()

# Append text to the third paragraph.
paragraph.AppendText("This is the third paragraph.")

# Create a document navigator to facilitate navigation and formatting.
navigator = DocumentNavigator(doc)

# Apply bullet list style to the current position (first paragraph by default).
navigator.ListFormat.ApplyBulletStyle()

# Move the navigator's cursor to the third paragraph (index 2) at character offset 0.
navigator.MoveToParagraph(2, 0)

# Apply bullet list style to the third paragraph.
navigator.ListFormat.ApplyBulletStyle()

# Save the document to a file named "ListFormat.docx" in DOCX format.
doc.SaveToFile("ListFormat.docx", FileFormat.Docx)

# Close the document to release internal resources.
doc.Close()

# Explicitly dispose of the document object to free memory immediately.
doc.Dispose()