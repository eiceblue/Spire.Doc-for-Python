from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Sample.docx"
outputFile = "DocumentProperty.docx"

#Open a blank Word document as template.
document = Document()
document.LoadFromFile(inputFile)
document.BuiltinDocumentProperties.Title = "Document Demo Document"
document.BuiltinDocumentProperties.Subject = "demo"
document.BuiltinDocumentProperties.Author = "James"
document.BuiltinDocumentProperties.Company = "e-iceblue"
document.BuiltinDocumentProperties.Manager = "Jakson"
document.BuiltinDocumentProperties.Category = "Doc Demos"
document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo"
document.BuiltinDocumentProperties.Comments = "This document is just a demo."
#Save as docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()