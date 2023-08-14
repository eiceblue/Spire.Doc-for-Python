import unittest
from spire.doc import *
from spire.doc.common import *

outputFile = "InsertFootnote.docx"
inputFile = "./Data/FootnoteExample.docx"

document = Document()
document.LoadFromFile(inputFile)

#finds the first matched string.
selection = document.FindString("Spire.Doc", False, True)
textRange = selection.GetAsOneRange()
paragraph = textRange.OwnerParagraph
index = paragraph.ChildObjects.IndexOf(textRange)
footnote = paragraph.AppendFootnote(FootnoteType.Footnote)
paragraph.ChildObjects.Insert(index + 1, footnote)
textRange = footnote.TextBody.AddParagraph().AppendText("Welcome to evaluate Spire.Doc")
textRange.CharacterFormat.FontName = "Arial Black"
textRange.CharacterFormat.FontSize = 10
textRange.CharacterFormat.TextColor = Color.get_DarkGray()
footnote.MarkerCharacterFormat.FontName = "Calibri"
footnote.MarkerCharacterFormat.FontSize = 12
footnote.MarkerCharacterFormat.Bold = True
footnote.MarkerCharacterFormat.TextColor = Color.get_DarkGreen()
document.SaveToFile(outputFile, FileFormat.Docx2010)
document.Close()