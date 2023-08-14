from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/ToRtf.doc"
outputFile = "ToRtf.rtf"
#Create word document
document = Document()
document.LoadFromFile(inputFile)
#Save doc file.
document.SaveToFile(outputFile, FileFormat.Rtf)
document.Close()
