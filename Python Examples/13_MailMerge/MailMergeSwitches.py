from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/MailMergeSwitches.docx"
outputFile = "MailMergeSwitches.docx"

doc = Document()
#Load a mail merge template file
doc.LoadFromFile(inputFile)

fieldName = ["XX_Name"]
fieldValue = ["Jason Tang"]

doc.MailMerge.Execute(fieldName, fieldValue)
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
