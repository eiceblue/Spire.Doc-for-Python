import unittest
from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/MailMerge.doc"
outputFile = "MailMerge.doc"

#Create word document
document = Document()
document.LoadFromFile(inputFile)

filedNames = ["Contact Name", "Fax", "Date"]

filedValues = ["John Smith", "+1 (69) 123456", DateTime.get_Now().Date.ToString()]

document.MailMerge.Execute(filedNames, filedValues)

#Save doc file.
document.SaveToFile(outputFile, FileFormat.Doc)
document.Close()