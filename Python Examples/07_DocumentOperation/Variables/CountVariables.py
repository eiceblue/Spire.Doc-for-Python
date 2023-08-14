from spire.doc import *
from spire.doc.common import *

def WriteAllText(fname:str,text:List[str]):
        fp = open(fname,"w")
        for s in text:
            fp.write(s)
        fp.close()

inputFile = "./Data/Template_Docx_6.docx"
outputFile = "CountVariables.txt"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Get the number of variables in the document.
number = document.Variables.Count
content = ""
content += "The number of variables is: "
content += str(number)
content += "\n"
#Save to file.
WriteAllText(outputFile, content)
document.Close()