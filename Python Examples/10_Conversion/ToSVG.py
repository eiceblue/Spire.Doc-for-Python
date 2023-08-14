from spire.doc import *
from spire.doc.common import *



inputFile = "./Data/ToSVGTemplate.docx"
outputFile = "ToSVG.svg"
#Create word document
document = Document()
document.LoadFromFile(inputFile)
document.SaveToFile(outputFile, FileFormat.SVG)
document.Close()