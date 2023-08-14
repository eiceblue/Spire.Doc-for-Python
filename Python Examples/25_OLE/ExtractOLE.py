from spire.doc import *
from spire.doc.common import *

outputFile_pdf = "ExtractOLE.pdf"
outputFile_xls = "ExtractOLE.xls"
outputFile_pptx = "ExtractOLE.pptx"
inputFile = "./Data/OLEs.docx"

#Create document and load file from disk
doc = Document()
doc.LoadFromFile(inputFile)

#Traverse through all sections of the word document    
for k in range(doc.Sections.Count):
    sec = doc.Sections.get_Item(k)
    #Traverse through all Child Objects in the body of each section
    for j in range(sec.Body.ChildObjects.Count):
        obj = sec.Body.ChildObjects.get_Item(j)
        #find the paragraph
        if isinstance(obj, Paragraph):
            par = obj if isinstance(obj, Paragraph) else None
            for m in range(par.ChildObjects.Count):
                o = par.ChildObjects.get_Item(m)
                #check whether the object is OLE
                if o.DocumentObjectType == DocumentObjectType.OleObject:
                    Ole = o if isinstance(o, DocOleObject) else None
                    s = Ole.ObjectType
                    #check whether the object type is "Acrobat.Document.11"
                    if s == "AcroExch.Document.DC":
                        #write the data of OLE into file
                        fp = open(outputFile_pdf,"wb")
                        fp.write(Ole.NativeData)
                        fp.close()

                    #check whether the object type is "Excel.Sheet.8"
                    elif s == "Excel.Sheet.8":
                        fp = open(outputFile_xls,"wb")
                        fp.write(Ole.NativeData)
                        fp.close()
                    #check whether the object type is "PowerPoint.Show.12"
                    elif s == "PowerPoint.Show.12":
                        fp = open(outputFile_pptx,"wb")
                        fp.write(Ole.NativeData)
                        fp.close()
doc.Close()