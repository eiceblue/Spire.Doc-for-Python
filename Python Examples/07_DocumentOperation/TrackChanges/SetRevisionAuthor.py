from spire.doc import *
from spire.doc.common import *

# Initialize a new document instance.
document = Document()
# Load the file from disk into the document instance.
document.LoadFromFile("Data/GetRevisions.docx")
# Loop through each section in the document.
for i in range(document.Sections.Count):
    # Get the current section by index.
    sec = document.Sections.get_Item(i)
    # Iterate through each element under the body of the section.
    for j in range(sec.Body.ChildObjects.Count):
        # Get the current item within the body objects by index.
        docItem = sec.Body.ChildObjects.get_Item(j)
        # Check if the current item is an instance of Paragraph.
        if isinstance(docItem, Paragraph):
            # Assign the paragraph to a variable for easy reference.
            para = docItem
            # Check if the paragraph is an insertion revision.
            if para.IsInsertRevision:  
                # Set the author of the insertion revision.
                para.InsertRevision.Author="E-iceblue"
            # Otherwise, check if the paragraph is a deletion revision.
            elif para.IsDeleteRevision:
                # Set the author of the deletion revision.
                para.DeleteRevision.Author="E-iceblue"
            # Iterate through each child object within the paragraph.
            for k in range(para.ChildObjects.Count):
                # Get the current text range item within paragraph's child objects.
                textRange = para.ChildObjects.get_Item(k)
                # Check if the current item is an instance of TextRange.
                if isinstance(textRange, TextRange):
                    # Check if the text range is an insertion revision.
                    if textRange.IsInsertRevision:
                        # Set the author of the insertion revision for text range.
                        textRange.InsertRevision.Author="E-iceblue"
                    # Otherwise, check if the text range is a deletion revision.
                    elif textRange.IsDeleteRevision:
                        # Set the author of the deletion revision for text range.
                        textRange.DeleteRevision.Author="E-iceblue"
# Save the modified document to a file with specified format.
document.SaveToFile("SetRevisionAuthor.docx", FileFormat.Docx2013)
# Close the document instance.
document.Close()