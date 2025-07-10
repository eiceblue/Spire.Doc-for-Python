from spire.doc import *
from spire.doc.common import *

inputFile = "data/Sample.docx"
outputFile = "SetDocumentProperties.docx"

#Create a document
document = Document()

#Load the document from disk.
document.LoadFromFile(inputFile)

#Set the build-in Properties.
document.BuiltinDocumentProperties.Title = "Document Demo Document"
document.BuiltinDocumentProperties.Author = "James"
document.BuiltinDocumentProperties.Company = "e-iceblue"
document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo"
document.BuiltinDocumentProperties.Comments = "This document is just a demo."

#Set the custom properties.
custom = document.CustomDocumentProperties
custom.Add("e-iceblue", Boolean(True))
custom.Add("Authorized By", String("John Smith"))
custom.Add("Authorized Date", DateTime.get_Today())

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
