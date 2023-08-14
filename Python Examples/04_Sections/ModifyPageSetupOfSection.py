from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_N2.docx"
outputFile = "ModifyPageSetupOfSection.docx"

#Load Word from disk
doc = Document()
doc.LoadFromFile(inputFile)

#Loop through all sections
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    #Modify the margins
    section.PageSetup.Margins = MarginsF(100.0, 80.0, 100.0, 80.0)
    #Modify the page size
    section.PageSetup.PageSize = PageSize.Letter()


#Save the Word file
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()