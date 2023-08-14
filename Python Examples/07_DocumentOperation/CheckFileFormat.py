from spire.doc import *
from spire.doc.common import *

def WriteAllText(fname:str,text:List[str]):
        fp = open(fname,"w")
        for s in text:
            fp.write(s)
        fp.close()

inputFile = "./Data/Template.docx"
outputFile = "CheckFileFormat.txt"

#Create Word document.
doc = Document()
doc.LoadFromFile(inputFile)
#Get file format
ff = doc.DetectedFormatType
fileFormat = "The file format is "
#Check the format info
if ff == FileFormat.Doc:
    fileFormat += "Microsoft Word 97-2003 document."
elif ff == FileFormat.Dot:
    fileFormat += "Microsoft Word 97-2003 template."
elif ff == FileFormat.Docx:
    fileFormat += "Office Open XML WordprocessingML Macro-Free Document."
elif ff == FileFormat.Docm:
    fileFormat += "Office Open XML WordprocessingML Macro-Enabled Document."
elif ff == FileFormat.Dotx:
    fileFormat += "Office Open XML WordprocessingML Macro-Free Template."
elif ff == FileFormat.Dotm:
    fileFormat += "Office Open XML WordprocessingML Macro-Enabled Template."
elif ff == FileFormat.Rtf:
    fileFormat += "RTF format."
elif ff == FileFormat.WordML:
    fileFormat += "Microsoft Word 2003 WordprocessingML format."
elif ff == FileFormat.Html:
    fileFormat += "HTML format."
elif ff == FileFormat.WordXml:
    fileFormat += "Microsoft Word xml format for word 2007-2013."
elif ff == FileFormat.Odt:
    fileFormat += "OpenDocument Text."
elif ff == FileFormat.Ott:
    fileFormat += "OpenDocument Text Template."
elif ff == FileFormat.DocPre97:
    fileFormat += "Microsoft Word 6 or Word 95 format."
else:
    fileFormat += "Unknown format."
#Save to file.
WriteAllText(outputFile, fileFormat)
doc.Close()
