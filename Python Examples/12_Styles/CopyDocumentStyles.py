import unittest
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Toc.docx"
inputFile2 = "./Data/Template_N3.docx"
outputFile = "CopyDocumentStyles.docx"

#Load source document from disk
srcDoc = Document()
srcDoc.LoadFromFile(inputFile)

#Load destination document from disk
destDoc = Document()
destDoc.LoadFromFile(inputFile2)

#Get the style collections of source document
styles = srcDoc.Styles

#Add the style to destination document
for i in range(styles.Count):
    style = styles.get_Item(i)
    destDoc.Styles.Add(style)

#Save the Word file
destDoc.SaveToFile(outputFile, FileFormat.Docx2013)
destDoc.Close()
