from spire.doc import *
from spire.doc.common import *

outputFile = "Macros.docm"
inputFile = "./Data/Macros.docm"

document = Document()
#Loading documetn with macros.
document.LoadFromFile(inputFile, FileFormat.Docm)

#Save docm file.
document.SaveToFile(outputFile, FileFormat.Docm)
document.Close()