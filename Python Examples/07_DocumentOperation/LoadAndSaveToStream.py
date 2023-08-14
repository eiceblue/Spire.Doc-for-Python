from spire.doc import *
from spire.doc.common import *

def WriteAllBytes(fname:str,data):
        fp = open(fname,"wb")
        #for d in data:
            #fp.write(d)
        fp.write(data)
        fp.close()

inputFile = "./Data/Sample.docx"
outputFile = "LoadAndSaveToStream.rtf"

# Open the stream. Read only access is enough to load a document.
stream = Stream(inputFile)
# Load the entire document into memory.
doc = Document(stream)
# You can close the stream now, it is no longer needed because the document is in memory.
stream.Close()
# Do something with the document
# Convert the document to a different format and save to stream.
newStream = Stream()
doc.SaveToStream(newStream, FileFormat.Rtf)
# Rewind the stream position back to zero so it is ready for the next reader.
newStream.Position = 0
newStream.Save(outputFile)
# Save the document from stream, to disk. Normally you would do something with the stream directly,
# For example, writing the data to a database.
WriteAllBytes(outputFile, newStream.ToArray())
doc.Close()