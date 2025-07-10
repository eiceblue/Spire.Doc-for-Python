from spire.doc import *
from spire.doc.common import *

# Create a Document object
doc = Document()

# Load a Word document
doc.LoadFromFile("Data/ToMarkdown.docx")

# Convert to Markdown format
doc.SaveToFile("ToMarkdown_output.md", FileFormat.Markdown)
