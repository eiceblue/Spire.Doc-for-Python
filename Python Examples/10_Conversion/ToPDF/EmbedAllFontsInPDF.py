from spire.doc import *
from spire.doc.common import *


inputFile =  "./Data/ConvertedTemplate.docx"
outputFile =  "EmbedAllFontsInPDF.pdf"
        
document = Document()
document.LoadFromFile(inputFile)
#embeds full fonts by default when IsEmbeddedAllFonts is set to true.
ppl = ToPdfParameterList()
ppl.IsEmbeddedAllFonts = True
#Save doc file to pdf.
document.SaveToFile(outputFile, ppl)
document.Close()