
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ToEpub.doc"
outputFile = "ToEpub.epub"
doc = Document()
doc.LoadFromFile(inputFile)
doc.SaveToFile(outputFile, FileFormat.EPub)
doc.Close()