from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Template.docx"
outputFile = "ConvertDocToByte.docx"
doc = Document()
# Load the document from disk.
doc.LoadFromFile(inputFile)
# Create a new memory stream.
outStream = Stream()
# Save the document to stream.
doc.SaveToStream(outStream, FileFormat.Docx)
# Convert the document to bytes.
docBytes = outStream.ToArray()
# The bytes are now ready to be stored/transmitted.
# Now reverse the steps to load the bytes back into a document object.
inStream = Stream(docBytes)
# Load the stream into a new document object.
newDoc = Document(inStream)
#save doc file.
ms = Stream()
newDoc.SaveToStream(ms, FileFormat.Docx)
fp = open(outputFile,"wb")
#for d in data:
    #fp.write(d)
fp.write(ms.ToArray())
fp.close()
#File.WriteAllBytes(outputFile, ms.ToArray())
newDoc.Close()
