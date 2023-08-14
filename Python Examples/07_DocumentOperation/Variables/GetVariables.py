from spire.doc import *
from spire.doc.common import *

def WriteAllText(fname:str,text:List[str]):
        fp = open(fname,"w")
        for s in text:
            fp.write(s)
        fp.close()

inputFile = "./Data/Template_Docx_6.docx"
outputFile = "GetVariables.txt"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
stringBuilder = ""
stringBuilder += "This document has following variables:"
stringBuilder += "\n"
for i in range(document.Variables.Count):
    name = document.Variables.GetNameByIndex(i)
    value = document.Variables.GetValueByIndex(i)
    stringBuilder += "Name: " 
    stringBuilder += name 
    stringBuilder += ", "
    stringBuilder += "Value: " 
    stringBuilder += value
    stringBuilder += "\n"
WriteAllText(outputFile, stringBuilder)
document.Close()