
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ReplaceTextByRegex.docx"
outputFile = "ReplaceTextByRegex.docx"

#Create the first document
document = Document()

#Load the document from disk.
document.LoadFromFile(inputFile)

#create a regex, match the text that starts with #
regex = Regex("""\\#\\w+\\b""")

#replace the text by regex
document.Replace(regex, "Spire.Doc")

#save the document
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
