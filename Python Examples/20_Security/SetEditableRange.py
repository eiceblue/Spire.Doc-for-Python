from spire.doc import *
from spire.doc.common import *


outputFile = "SetEditableRange.docx"
inputFile = "./Data/SetEditableRange.docx"


# Create a new document
document = Document()
# Load file from disk
document.LoadFromFile(inputFile)
# Protect whole document
document.Protect(ProtectionType.AllowOnlyReading, "password")
# Create tags for permission start and end
start = PermissionStart(document, "testID")
end = PermissionEnd(document, "testID")
# Add the start and end tags to allow the first paragraph to be edited.
document.Sections[0].Paragraphs[0].ChildObjects.Insert(0, start)
document.Sections[0].Paragraphs[0].ChildObjects.Add(end)
# Save the document
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
