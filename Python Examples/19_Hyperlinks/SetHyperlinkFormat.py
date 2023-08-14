from spire.doc import *
from spire.doc.common import *

outputFile = "SetHyperlinkFormat.docx"
inputFile = "./Data/BlankTemplate.docx"

# Load Document
doc = Document()
doc.LoadFromFile(inputFile)
section = doc.Sections[0]

# Add a paragraph and append a hyperlink to the paragraph
para1 = section.AddParagraph()
para1.AppendText("Regular Link: ")
# Format the hyperlink with default color and underline style
txtRange1 = para1.AppendHyperlink(
    "www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink)
txtRange1.CharacterFormat.FontName = "Times New Roman"
txtRange1.CharacterFormat.FontSize = 12
blankPara1 = section.AddParagraph()

# Add a paragraph and append a hyperlink to the paragraph
para2 = section.AddParagraph()
para2.AppendText("Change Color: ")
# Format the hyperlink with red color and underline style
txtRange2 = para2.AppendHyperlink(
    "www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink)
txtRange2.CharacterFormat.FontName = "Times New Roman"
txtRange2.CharacterFormat.FontSize = 12
txtRange2.CharacterFormat.TextColor = Color.get_Red()
blankPara2 = section.AddParagraph()

# Add a paragraph and append a hyperlink to the paragraph
para3 = section.AddParagraph()
para3.AppendText("Remove Underline: ")
# Format the hyperlink with red color and no underline style
txtRange3 = para3.AppendHyperlink(
    "www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink)
txtRange3.CharacterFormat.FontName = "Times New Roman"
txtRange3.CharacterFormat.FontSize = 12
txtRange3.CharacterFormat.UnderlineStyle = UnderlineStyle.none

# Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
