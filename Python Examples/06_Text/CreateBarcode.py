from spire.doc import *
from spire.doc.common import *

outputFile = "CreateBarcode.docx"

#Create a document
doc = Document()

#Add a paragraph
p = doc.AddSection().AddParagraph()

#Add barcode and set its format
txtRang = p.AppendText("H63TWX11072")
#Set barcode font name, note you need to install the barcode font on your system at first
txtRang.CharacterFormat.FontName = "C39HrP60DlTt"
txtRang.CharacterFormat.FontSize = 80
txtRang.CharacterFormat.TextColor = Color.get_SeaGreen()

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()

