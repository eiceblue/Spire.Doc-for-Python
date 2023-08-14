from spire.doc import *
from spire.doc.common import *

def WriteAllText(fname:str,text:List[str]):
        fp = open(fname,"w")
        for s in text:
            fp.write(s)
        fp.close()

inputFile = "./Data/Template_Docx_1.docx"
outputFile =  "GetText.txt"
     
#Load the document from disk.
document = Document()
document.LoadFromFile(inputFile)

#get text from document
text = document.GetText()

#create a new TXT File to save extracted text
WriteAllText(outputFile, text)
document.Close()