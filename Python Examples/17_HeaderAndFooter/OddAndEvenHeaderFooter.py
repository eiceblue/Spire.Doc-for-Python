from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/MultiplePages.docx"
outputFile = "OddAndEvenHeaderFooter.docx"

# Load the document
doc = Document()
doc.LoadFromFile(inputFile)

# Get the section and
section = doc.Sections[0]

# Set the DifferentOddAndEvenPagesHeaderFooter property to ture
section.PageSetup.DifferentOddAndEvenPagesHeaderFooter = True

# Add odd header
P3 = section.HeadersFooters.OddHeader.AddParagraph()
OH = P3.AppendText("Odd Header")
P3.Format.HorizontalAlignment = HorizontalAlignment.Center
OH.CharacterFormat.FontName = "Arial"
OH.CharacterFormat.FontSize = 10

# Add even header
P4 = section.HeadersFooters.EvenHeader.AddParagraph()
EH = P4.AppendText("Even Header from E-iceblue Using Spire.Doc")
P4.Format.HorizontalAlignment = HorizontalAlignment.Center
EH.CharacterFormat.FontName = "Arial"
EH.CharacterFormat.FontSize = 10

# Add odd footer
P2 = section.HeadersFooters.OddFooter.AddParagraph()
OF = P2.AppendText("Odd Footer")
P2.Format.HorizontalAlignment = HorizontalAlignment.Center
OF.CharacterFormat.FontName = "Arial"
OF.CharacterFormat.FontSize = 10

# Add even footer
P1 = section.HeadersFooters.EvenFooter.AddParagraph()
EF = P1.AppendText("Even Footer from E-iceblue Using Spire.Doc")
EF.CharacterFormat.FontName = "Arial"
EF.CharacterFormat.FontSize = 10
P1.Format.HorizontalAlignment = HorizontalAlignment.Center

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
