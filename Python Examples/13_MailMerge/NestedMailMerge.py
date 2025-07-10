from spire.doc import *
from spire.doc.common import *

outputFile = "Demo/NestedMailMerage.docx"
inputFile = "Data/NestedMailMerge.doc"

#Create word document
document = Document()
document.LoadFromFile(inputFile)

# execute mailmerge
tempdDict = {"Customer": '', "Order": "Customer_Id = %Customer.Customer_Id%"}
dataFile = "Data/Orders.xml"
document.MailMerge.ExecuteWidthNestedRegion(dataFile, tempdDict)

#Save as docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
