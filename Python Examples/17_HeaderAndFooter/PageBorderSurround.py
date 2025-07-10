from spire.doc import *
from spire.doc.common import *

outputFile = "PageBorderSurround.docx"

# Create a new document
doc = Document()
section = doc.AddSection()

# Add a sample page border to the document
section.PageSetup.Borders.BorderType=BorderStyle.Wave
section.PageSetup.Borders.Color=Color.get_Green()
section.PageSetup.Borders.Left.Space = 20.0
section.PageSetup.Borders.Right.Space = 20.0

# Add a header and set its format
paragraph1 = section.HeadersFooters.Header.AddParagraph()
paragraph1.Format.HorizontalAlignment = HorizontalAlignment.Right
headerText = paragraph1.AppendText("Header isn't included in page border")
headerText.CharacterFormat.FontName = "Calibri"
headerText.CharacterFormat.FontSize = 20.0
headerText.CharacterFormat.Bold = True

# Add a footer and set its format
paragraph2 = section.HeadersFooters.Footer.AddParagraph()
paragraph2.Format.HorizontalAlignment = HorizontalAlignment.Left
footerText = paragraph2.AppendText("Footer is included in page border")
footerText.CharacterFormat.FontName = "Calibri"
footerText.CharacterFormat.FontSize = 20.0
footerText.CharacterFormat.Bold = True

# Set the header not included in the page border while the footer included
section.PageSetup.PageBorderIncludeHeader = False
section.PageSetup.HeaderDistance = 40.0
section.PageSetup.PageBorderIncludeFooter = True
section.PageSetup.FooterDistance = 40.0

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
