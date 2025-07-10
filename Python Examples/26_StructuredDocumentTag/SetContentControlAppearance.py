from spire.doc import *
from spire.doc.common import *

inputFile = "Data/ContentControl.docx"
outputFile = "SetContentControlAppearance.docx"

# Create a new document object
doc = Document()

# Load a document from the specified input file
doc.LoadFromFile(inputFile)

# Iterate through the sections in the document
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    # Iterate through the child objects in the section's body
    for j in range(section.Body.ChildObjects.Count):
        docObj = section.Body.ChildObjects.get_Item(j)
        # Check if the current object is a StructureDocumentTag
        if isinstance(docObj, StructureDocumentTag):
            # Get the StructureDocumentTag object and its SDTProperties
            stdTag = docObj
            sDTProperties = stdTag.SDTProperties

            # Set the appearance of the StructureDocumentTag based on its SDTType
            if sDTProperties.SDTType == SdtType.Text:
                sDTProperties.Appearance = SdtAppearance.BoundingBox
            elif sDTProperties.SDTType == SdtType.RichText:
                sDTProperties.Appearance = SdtAppearance.Hidden
            elif sDTProperties.SDTType == SdtType.Picture:
                sDTProperties.Appearance = SdtAppearance.Tags
            elif sDTProperties.SDTType == SdtType.CheckBox:
                sDTProperties.Appearance = SdtAppearance.Default

# Save the modified document to the output file
doc.SaveToFile(outputFile, FileFormat.Docx2013)

# Dispose the document object
doc.Dispose()
