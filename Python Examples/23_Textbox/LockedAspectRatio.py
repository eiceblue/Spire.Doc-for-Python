# 未生效 textBox1.AspectRatioLocked = True
from spire.doc import *
from spire.doc.common import *

outputFile = "LockedAspectRatio.docx"

# Create a new instance of Document
document = Document()

# Add a new section to the document
section = document.AddSection()

# Add a paragraph to the section
paragraph = section.AddParagraph()

# Append a textbox to the paragraph and get a reference to it
textBox1 = paragraph.AppendTextBox(240, 35)

# Configure the horizontal alignment, line color, and line style of the textbox
textBox1.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left
textBox1.Format.LineColor = Color.get_Black()
textBox1.Format.LineStyle = TextBoxLineStyle.Simple

# Lock the aspect ratio of the textbox
textBox1.AspectRatioLocked = True

# Add a paragraph to the body of the textbox and get a reference to it
para = textBox1.Body.AddParagraph()

# Add text to the paragraph
txtrg = para.AppendText("Textbox 1 in the document")
txtrg.CharacterFormat.FontName = "Lucida Sans Unicode"
txtrg.CharacterFormat.FontSize = 14
txtrg.CharacterFormat.TextColor = Color.get_Black()
para.Format.HorizontalAlignment = HorizontalAlignment.Center

# Save the document to a file named "Sample.docx" in Docx format
document.SaveToFile(outputFile, FileFormat.Docx)

# Dispose the document object to release resources
document.Dispose()
