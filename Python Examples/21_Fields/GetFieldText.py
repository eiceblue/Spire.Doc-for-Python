from spire.doc import *
from spire.doc.common import *


def WriteAllText(fpath: str, content: str):
    with open(fpath, 'w') as fp:
        fp.write(content)


outputFile = "GetFieldText.txt"
inputFile = "./Data/SampleB_1.docx"

sb = ''

# Open a Word document
document = Document()
document.LoadFromFile(inputFile)

# Get all fields in document
fields = document.Fields

for i in range(fields.Count):
    field = fields.get_Item(i)
    # Get field text
    fieldText = field.FieldText
    sb += "The field text is \"" + fieldText + "\".\r\n"
    sb += "\n"

WriteAllText(outputFile, sb)
document.Close()
