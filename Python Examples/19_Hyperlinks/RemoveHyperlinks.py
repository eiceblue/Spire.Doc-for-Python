from spire.doc import *
from spire.doc.common import *


def _FindAllHyperlinks(document):
    hyperlinks = []
    # Iterate through the items in the sections to find all hyperlinks
    for i in range(document.Sections.Count):
        section = document.Sections.get_Item(i)
        for j in range(section.Body.ChildObjects.Count):
            sec = section.Body.ChildObjects.get_Item(j)
            if sec.DocumentObjectType == DocumentObjectType.Paragraph:
                for k in range((sec if isinstance(sec, Paragraph) else None).ChildObjects.Count):
                    para = (sec if isinstance(sec, Paragraph)
                            else None).ChildObjects.get_Item(k)
                    if para.DocumentObjectType == DocumentObjectType.Field:
                        field = para if isinstance(para, Field) else None
                        if field.Type == FieldType.FieldHyperlink:
                            hyperlinks.append(field)
    return hyperlinks

# Flatten the hyperlink field
def _FlattenHyperlinks(field):
    ownerParaIndex = field.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(
        field.OwnerParagraph)
    fieldIndex = field.OwnerParagraph.ChildObjects.IndexOf(field)
    sepOwnerPara = field.Separator.OwnerParagraph
    sepOwnerParaIndex = field.Separator.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(
        field.Separator.OwnerParagraph)
    sepIndex = field.Separator.OwnerParagraph.ChildObjects.IndexOf(
        field.Separator)
    endIndex = field.End.OwnerParagraph.ChildObjects.IndexOf(field.End)
    endOwnerParaIndex = field.End.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(
        field.End.OwnerParagraph)

    _FormatFieldResultText(field.Separator.OwnerParagraph.OwnerTextBody,
                           sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex)

    field.End.OwnerParagraph.ChildObjects.RemoveAt(endIndex)

    for i in range(sepOwnerParaIndex, ownerParaIndex - 1, -1):
        if i == sepOwnerParaIndex and i == ownerParaIndex:
            for j in range(sepIndex, fieldIndex - 1, -1):
                field.OwnerParagraph.ChildObjects.RemoveAt(j)

        elif i == ownerParaIndex:
            for j in range(field.OwnerParagraph.ChildObjects.Count - 1, fieldIndex - 1, -1):
                field.OwnerParagraph.ChildObjects.RemoveAt(j)

        elif i == sepOwnerParaIndex:
            for j in range(sepIndex, -1, -1):
                sepOwnerPara.ChildObjects.RemoveAt(j)
        else:
            field.OwnerParagraph.OwnerTextBody.ChildObjects.RemoveAt(i)

# Remove the font color and underline format of the hyperlinks
def _FormatFieldResultText(ownerBody, sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex):
    for i in range(sepOwnerParaIndex, endOwnerParaIndex + 1):
        para = ownerBody.ChildObjects[i] if isinstance(
            ownerBody.ChildObjects[i], Paragraph) else None
        if i == sepOwnerParaIndex and i == endOwnerParaIndex:
            for j in range(sepIndex + 1, endIndex):
                _FormatText(para.ChildObjects[j] if isinstance(
                    para.ChildObjects[j], TextRange) else None)

        elif i == sepOwnerParaIndex:
            for j in range(sepIndex + 1, para.ChildObjects.Count):
                _FormatText(para.ChildObjects[j] if isinstance(
                    para.ChildObjects[j], TextRange) else None)
        elif i == endOwnerParaIndex:
            for j in range(0, endIndex):
                _FormatText(para.ChildObjects[j] if isinstance(
                    para.ChildObjects[j], TextRange) else None)
        else:
            for j, unusedItem in enumerate(para.ChildObjects):
                _FormatText(para.ChildObjects[j] if isinstance(
                    para.ChildObjects[j], TextRange) else None)


def _FormatText(tr):
    # Set the text color to black
    tr.CharacterFormat.TextColor = Color.get_Black()
    # Set the text underline style to none
    tr.CharacterFormat.UnderlineStyle = UnderlineStyle.none


outputFile = "RemoveHyperlinks.docx"
inputFile = "./Data/Hyperlinks.docx"

# Load Document
doc = Document()
doc.LoadFromFile(inputFile)

# Get all hyperlinks
hyperlinks = _FindAllHyperlinks(doc)

# Flatten all hyperlinks
for i in range(len(hyperlinks) - 1, -1, -1):
    _FlattenHyperlinks(hyperlinks[i])

    # Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
