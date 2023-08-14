from spire.doc import *
from spire.doc.common import *

outputFile = "RemoveEditableRange.docx"
inputFile = "./Data/RemoveEditableRange.docx"

# Create a new document
document = Document()
# Load file from disk
document.LoadFromFile(inputFile)
# Find "PermissionStart" and "PermissionEnd" tags and remove them
for k in range(document.Sections.Count):
    section = document.Sections.get_Item(k)
    for j in range(section.Body.Paragraphs.Count):
        paragraph = section.Body.Paragraphs.get_Item(j)
        i = 0
        while i < paragraph.ChildObjects.Count:
            obj = paragraph.ChildObjects[i]
            if isinstance(obj, PermissionStart) or isinstance(obj, PermissionEnd):
                paragraph.ChildObjects.Remove(obj)
            else:
                i += 1

# Save to file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
