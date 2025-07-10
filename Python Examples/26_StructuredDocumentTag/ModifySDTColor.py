### 颜色给失败

from spire.doc import *
from spire.doc.common import *

inputFile = "Data/ModifySTDColor.docx"
outputFile = "ModifySTDColor.docx"

# Create a new document object
doc = Document()

# Load a document from the specified file path
doc.LoadFromFile(inputFile)

# Iterate through the sections in the document
for s in range(doc.Sections.Count):
    # Get the current section
    section = doc.Sections[s]

    # Iterate through the child objects in the section's body
    for i in range(section.Body.ChildObjects.Count):
        # Check if the child object is a Paragraph
        if isinstance(section.Body.ChildObjects[i], Paragraph):
            # Get the paragraph object
            para = section.Body.ChildObjects[i] if isinstance(
                section.Body.ChildObjects[i], Paragraph) else None

            # Iterate through the child objects in the paragraph
            for j in range(para.ChildObjects.Count):
                # Check if the child object is a StructureDocumentTagInline
                if isinstance(para.ChildObjects[j],
                              StructureDocumentTagInline):
                    # Get the StructureDocumentTagInline object
                    sdt = para.ChildObjects[j] if isinstance(
                        para.ChildObjects[j],
                        StructureDocumentTagInline) else None

                    # Get the SDTProperties of the StructureDocumentTagInline
                    sDTProperties = sdt.SDTProperties

                    # Set the color of the SDTProperties based on the SDTType
                    if sDTProperties.SDTType == SdtType.RichText:
                        print(1)
                        sDTProperties.Color = Color.get_Orange()
                    elif sDTProperties.SDTType == SdtType.Text:
                        print(2)
                        sDTProperties.Color = Color.get_Green()

        # Check if the child object is a StructureDocumentTag
        if isinstance(section.Body.ChildObjects[i], StructureDocumentTag):
            # Get the StructureDocumentTag object
            sdt = section.Body.ChildObjects[i] if isinstance(
                section.Body.ChildObjects[i], StructureDocumentTag) else None

            # Get the SDTProperties of the StructureDocumentTag
            sDTProperties = sdt.SDTProperties

            # Set the color of the SDTProperties based on the SDTType
            if sDTProperties.SDTType == SdtType.RichText:
                sDTProperties.Color = Color.Orange
            elif sDTProperties.SDTType == SdtType.Text:
                sDTProperties.Color = Color.Green
# Save the modified document to the output file in DOCX format
doc.SaveToFile(outputFile, FileFormat.Docx2013)

# Dispose the document object
doc.Dispose()
