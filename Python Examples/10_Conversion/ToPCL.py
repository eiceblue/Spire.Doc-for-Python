from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/ConvertedTemplate.docx"
outputFile = "ToPCL.pcl"
doc = Document()
doc.LoadFromFile(inputFile)
doc.SaveToFile(outputFile, FileFormat.PCL)
doc.Close()
