from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/HideEmptyRegions.doc"
outputFile = "HideEmptyRegions.docx"
#Create word document
document = Document()
document.LoadFromFile(inputFile)
filedNames = ["Contact Name", "Fax", "Date"]
filedValues = ["John Smith", "+1 (69) 123456", DateTime.get_Now().Date.ToString()]
#Set the value to remove paragraphs which contain empty field.
document.MailMerge.HideEmptyParagraphs = True
#Set the value to remove group which contain empty field.
document.MailMerge.HideEmptyGroup = True
document.MailMerge.Execute(filedNames, filedValues)
#Save doc file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
