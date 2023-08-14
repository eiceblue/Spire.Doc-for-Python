from spire.doc import *
from spire.doc.common import *


def WriteAllText(fpath: str, content: str):
    with open(fpath, 'w') as fp:
        fp.write(content)


outputFile = "GetMergeFieldName.txt"
inputFile = "./Data/MailMerge.doc"

sb = ''

# Open a Word document
document = Document()
document.LoadFromFile(inputFile)

# Get merge field name
fieldNames = document.MailMerge.GetMergeFieldNames()

sb = "The document has " + str(len(fieldNames)) + " merge fields."
sb += " The below is the name of the merge field:" + "\r\n"
for name in fieldNames:
    sb += name
    sb += '\n'

WriteAllText(outputFile, sb)
document.Close()
