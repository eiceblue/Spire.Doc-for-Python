from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ConvertedTemplate2.docx"
outputFile = "SpecifyEmbeddedFont.pdf"

document = Document()
document.LoadFromFile(inputFile)
#Specify embedded font
parms = ToPdfParameterList()
part = []
part.append("PT Serif Caption")
parms.EmbeddedFontNameList = part
document.SaveToFile(outputFile, parms)
document.Close()