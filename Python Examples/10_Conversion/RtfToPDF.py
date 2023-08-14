from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template_RtfFile.rtf"
outputFile = "RtfToPDF.pdf"

doc = Document()
doc.SetDateTimeOfUnitTest(DateTime.Parse("2022/5/1 00:00:00"))
doc.LoadFromFile(inputFile)

doc.SaveToFile(outputFile, FileFormat.PDF)
doc.Close()
