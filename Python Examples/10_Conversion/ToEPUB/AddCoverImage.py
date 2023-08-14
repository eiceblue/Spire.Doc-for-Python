
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ToEpub.doc"
outputFile = "AddCoverImage.epub"

doc = Document()
doc.LoadFromFile(inputFile)
picture = DocPicture(doc)
picture.LoadImage("./Data/Cover.png")
doc.SaveToEpub(outputFile, picture)
doc.Close()

