from spire.doc import *
from spire.doc.common import *

def WriteAllText(fname:str,text:List[str]):
        fp = open(fname,"w")
        for s in text:
            fp.write(s)
        fp.close()

inputFile = "./Data/Template_Docx_6.docx"
outputFile = "RetrieveVariables.txt"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Retrieve name of the variable by index.
s1 = document.Variables.GetNameByIndex(0)
#Retrieve value of the variable by index.
s2 = document.Variables.GetValueByIndex(0)
#Retrieve the value of the variable by name.
s3 = document.Variables["A1"]
content = ""
content += "The name of the variable retrieved by index 0 is: "
content += s1
content += "\n"
content += "The vaule of the variable retrieved by index 0 is: " 
content += s2
content += "\n"
content += "The vaule of the variable retrieved by name \"A1\" is: " 
content += s3
content += "\n"
#Save to file.
WriteAllText(outputFile, content)
document.Close()