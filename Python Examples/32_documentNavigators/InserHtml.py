from spire.doc import *

# Create a new Document object to represent an empty Word document.
doc = Document()

# Initialize a DocumentNavigator instance to simplify content insertion and navigation within the document.
navigator = DocumentNavigator(doc)

html = "<body style=\"font-family: Arial, sans-serif background-color: #f4f4f9 color: #333 padding: 20px line-height: 1.6\">\r\n    <h1 style=\"color: #2c3e50\">Welcome to the Random English Page!</h1>\r\n    <p>This is a randomly generated HTML document containing English text.</p>\r\n    <p>The quick brown <span style=\"background-color: #fffacd padding: 4px 8px border-radius: 4px\">fox jumps</span> over the lazy dog. This sentence contains every letter of the English alphabet.</p>\r\n    <p>Here are a few fun facts:</p>\r\n    <ul>\r\n        <li>English is spoken by over 1.5 billion people worldwide.</li>\r\n        <li>The word \"set\" has the most definitions in the English language.</li>\r\n        <li>\"Dreamt\" is the only English word that ends with \"mt\".</li>\r\n    </ul>\r\n    <p>Thank you for visiting! Have a wonderful day.</p>\r\n</body>"


# Insert the HTML content (currently empty) into the document at the current cursor position.
navigator.InsertHtml(html)

# Save the resulting document to a file in DOCX format.
doc.SaveToFile("InserHtml.docx", FileFormat.Docx)

# Close the document to release internal resources.
doc.Close()

# Explicitly dispose of the document object to free memory immediately.
doc.Dispose()