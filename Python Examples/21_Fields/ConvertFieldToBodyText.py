from spire.doc import *
from spire.doc.common import *

outputFile = "ConvertFieldToBodyText.docx"
inputFile = "./Data/TextInputField.docx"

# Create the source document
sourceDocument = Document()

# Load the source document from disk.
sourceDocument.LoadFromFile(inputFile)

# Traverse FormFields
for j in range(sourceDocument.Sections[0].Body.FormFields.Count):
    field = sourceDocument.Sections[0].Body.FormFields.get_Item(j)
    # Find FieldFormTextInput type field
    if field.Type == FieldType.FieldFormTextInput:
        # Get the paragraph
        paragraph = field.OwnerParagraph

        # Define variables
        startIndex = 0
        endIndex = 0

        # Create a new TextRange
        textRange = TextRange(sourceDocument)

        # Set text for textRange
        textRange.Text = paragraph.Text

        # Traverse DocumentObjectS of field paragraph
        for k in range(paragraph.ChildObjects.Count):
            obj = paragraph.ChildObjects.get_Item(k)
            # If its DocumentObjectType is BookmarkStart
            if obj.DocumentObjectType == DocumentObjectType.BookmarkStart:
                # Get the index
                startIndex = paragraph.ChildObjects.IndexOf(obj)
            # If its DocumentObjectType is BookmarkEnd
            if obj.DocumentObjectType == DocumentObjectType.BookmarkEnd:
                # Get the index
                endIndex = paragraph.ChildObjects.IndexOf(obj)
        # Remove ChildObjects
        for i in range(endIndex, startIndex, -1):
            # If it is TextFormField
            if isinstance(paragraph.ChildObjects[i], TextFormField):
                textFormField = paragraph.ChildObjects[i] if isinstance(
                    paragraph.ChildObjects[i], TextFormField) else None

                # Remove the field object
                paragraph.ChildObjects.Remove(textFormField)
            else:
                paragraph.ChildObjects.RemoveAt(i)
        # Insert the new TextRange
        paragraph.ChildObjects.Insert(startIndex, textRange)
        break

# Save the document.
sourceDocument.SaveToFile(outputFile, FileFormat.Docx)
sourceDocument.Close()
