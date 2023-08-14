from spire.doc import *
from spire.doc.common import *

outputFile = "Textbox.docx"

#Create a Word document and and a section.
document = Document()
section = document.AddSection()
paragraph = section.Paragraphs[0] if section.Paragraphs.Count > 0 else section.AddParagraph()
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()

#Insert and format the first textbox.
textBox1 = paragraph.AppendTextBox(240, 35)
textBox1.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left
textBox1.Format.LineColor = Color.get_Gray()
textBox1.Format.LineStyle = TextBoxLineStyle.Simple
textBox1.Format.FillColor = Color.get_DarkSeaGreen()
para = textBox1.Body.AddParagraph()
txtrg = para.AppendText("Textbox 1 in the document")
txtrg.CharacterFormat.FontName = "Lucida Sans Unicode"
txtrg.CharacterFormat.FontSize = 14
txtrg.CharacterFormat.TextColor = Color.get_White()
para.Format.HorizontalAlignment = HorizontalAlignment.Center

#Insert and format the second textbox.
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()
textBox2 = paragraph.AppendTextBox(240, 35)
textBox2.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left
textBox2.Format.LineColor = Color.get_Tomato()
textBox2.Format.LineStyle = TextBoxLineStyle.ThinThick
textBox2.Format.FillColor = Color.get_Blue()
textBox2.Format.LineDashing = LineDashing.Dot
para = textBox2.Body.AddParagraph()
txtrg = para.AppendText("Textbox 2 in the document")
txtrg.CharacterFormat.FontName = "Lucida Sans Unicode"
txtrg.CharacterFormat.FontSize = 14
txtrg.CharacterFormat.TextColor = Color.get_Pink()
para.Format.HorizontalAlignment = HorizontalAlignment.Center

#Insert and format the third textbox.
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()
textBox3 = paragraph.AppendTextBox(240, 35)
textBox3.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left
textBox3.Format.LineColor = Color.get_Violet()
textBox3.Format.LineStyle = TextBoxLineStyle.Triple
textBox3.Format.FillColor = Color.get_Pink()
textBox3.Format.LineDashing = LineDashing.DashDotDot
para = textBox3.Body.AddParagraph()
txtrg = para.AppendText("Textbox 3 in the document")
txtrg.CharacterFormat.FontName = "Lucida Sans Unicode"
txtrg.CharacterFormat.FontSize = 14
txtrg.CharacterFormat.TextColor = Color.get_Tomato()
para.Format.HorizontalAlignment = HorizontalAlignment.Center

#Save docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

        