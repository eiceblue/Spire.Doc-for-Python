
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ToOdt.doc"
outputFile = "ToOdt.odt"
#Create word document
document = Document()
document.LoadFromFile(inputFile)
#Save doc file.
document.SaveToFile(outputFile, FileFormat.Odt)
document.Close()
