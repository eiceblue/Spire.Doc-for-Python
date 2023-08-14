from spire.doc import *
from spire.doc.common import *

outputFile = "SetSuperscriptAndSubscript.docx"

#Create word document
document = Document()
#Create a new section
section = document.AddSection()
paragraph = section.AddParagraph()
paragraph.AppendText("E = mc")
range1 = paragraph.AppendText("2")
#Set supperscript
range1.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript
paragraph.AppendBreak(BreakType.LineBreak)
paragraph.AppendText("F")
range2 = paragraph.AppendText("n")
#Set subscript
range2.CharacterFormat.SubSuperScript = SubSuperScript.SubScript
paragraph.AppendText(" = F")
paragraph.AppendText("n-1").CharacterFormat.SubSuperScript = SubSuperScript.SubScript
paragraph.AppendText(" + F")
paragraph.AppendText("n-2").CharacterFormat.SubSuperScript = SubSuperScript.SubScript
#Set font size
for i in range(paragraph.Items.Count):
    item = paragraph.Items.get_Item(i)
    if isinstance(item, TextRange):
        ( item if isinstance(item, TextRange) else None).CharacterFormat.FontSize = 36     
#Save the document
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()