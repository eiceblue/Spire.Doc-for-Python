from spire.doc import *
from spire.doc.common import *

inputFile = "data/Revisions.docx"
outputFile = "ModifyRevisionTime.docx"

# Create a new Document object
document = Document()

# Load a Word document from a file
document.LoadFromFile(inputFile)

# Initialize index variables
index_insertRevision = 0
index_deleteRevision = 0

# Specify the date string and format
dateString = "2023/3/1 00:00:00"
formatStr = "yyyy/M/d HH:mm:ss"

# Parse the date string into a DateTime object using the specified format
date = DateTime.ParseExact(dateString, formatStr)

# Iterate through the sections in the document
for i in range(document.Sections.Count):
    sec = document.Sections[i]
    # Iterate through the child objects in the section's body
    for j in range(sec.Body.ChildObjects.Count):
        docItem = sec.Body.ChildObjects.get_Item(j)
        # Check if the child object is a Paragraph
        if isinstance(docItem, Paragraph):
            # Cast the child object to a Paragraph
            para = docItem

            # Check if the paragraph contains an insert revision
            if para.IsInsertRevision:
                # Increment the insert revision index
                index_insertRevision += 1

                # Get the InsertRevision object for the paragraph
                insRevison = para.InsertRevision

                # Set the DateTime property of the insert revision to the specified date
                insRevison.DateTime = date
            # Check if the paragraph contains a delete revision
            elif para.IsDeleteRevision:
                # Increment the delete revision index
                index_deleteRevision += 1

                # Get the DeleteRevision object for the paragraph
                delRevison = para.DeleteRevision

                # Set the DateTime property of the delete revision to the specified date
                delRevison.DateTime = date

            # Iterate through the child objects in the paragraph
            for k in range(para.ChildObjects.Count):
                obj = para.ChildObjects.get_Item(k)
                # Check if the child object is a TextRange
                if isinstance(obj, TextRange):
                    # Cast the child object to a TextRange
                    textRange = obj

                    # Check if the text range contains an insert revision
                    if textRange.IsInsertRevision:
                        # Increment the insert revision index
                        index_insertRevision += 1

                        # Get the InsertRevision object for the text range
                        insRevison = textRange.InsertRevision

                        # Set the DateTime property of the insert revision to the specified date
                        insRevison.DateTime = date
                    # Check if the text range contains a delete revision
                    elif textRange.IsDeleteRevision:
                        # Increment the delete revision index
                        index_deleteRevision += 1

                        # Get the DeleteRevision object for the text range
                        delRevison = textRange.DeleteRevision

                        # Set the DateTime property of the delete revision to the specified date
                        delRevison.DateTime = date

# Save the modified document to a new file
document.SaveToFile(outputFile, FileFormat.Docx)

# Dispose the Document object
document.Dispose()
