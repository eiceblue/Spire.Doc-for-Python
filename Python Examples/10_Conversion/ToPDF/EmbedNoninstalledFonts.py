from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/ConvertedTemplate2.docx"
outputFile = "EmbedNoninstalledFonts.pdf"

document = Document()
document.LoadFromFile(inputFile)

#Embed the non-installed fonts.
parms = ToPdfParameterList()
fonts = []
fonts.append(PrivateFontPath("PT Serif Caption", "./Data/PT_Serif-Caption-Web-Regular.ttf"))
parms.PrivateFontPaths = fonts

#Save doc file to pdf.
document.SaveToFile(outputFile, parms)
document.Close()
