from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ConvertedTemplate.docx"
outputFile = "ToPostScript.ps"
       
doc = Document()
doc.LoadFromFile(inputFile)
doc.SaveToFile(outputFile, FileFormat.PostScript)
doc.Close()
