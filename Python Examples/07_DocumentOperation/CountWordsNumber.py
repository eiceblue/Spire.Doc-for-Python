from spire.doc import *
from spire.doc.common import *

def WriteAllText(fname:str,text:List[str]):
        fp = open(fname,"w")
        for s in text:
            fp.write(s)
        fp.close()

inputFile = "./Data/Sample.docx"
outputFile = "CountWordsNumber.txt"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Count the number of words.
content = ""
content += "CharCount: " 
content += str(document.BuiltinDocumentProperties.CharCount)
content += "\n"
content += "CharCountWithSpace: " 
content += str(document.BuiltinDocumentProperties.CharCountWithSpace)
content += "\n"
content += "WordCount: " 
content += str(document.BuiltinDocumentProperties.WordCount)
content += "\n"
#Save to file.
WriteAllText(outputFile, content)
document.Close()