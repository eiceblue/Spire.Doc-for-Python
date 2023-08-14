from spire.doc import *
from spire.doc.common import *


def addSections(doc):
    #Add a section
    doc.AddSection()

def deleteSections(doc):
    #Delete the last section
    doc.Sections.RemoveAt(doc.Sections.Count - 1)

inputFile = "./Data/SectionTemplate.docx"
outputFile = "AddAndDeleteSections.docx"

doc = Document()
doc.LoadFromFile(inputFile)

addSections(doc)
deleteSections(doc)

doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()
