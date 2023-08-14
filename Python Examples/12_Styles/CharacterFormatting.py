import unittest
from spire.doc import *
from spire.doc.common import *

outputFile = "CharacterFormatting.docx"
#Initialize a document
document = Document()
sec = document.AddSection()
titleParagraph = sec.AddParagraph()
titleParagraph.AppendText("Font Styles and Effects ")
titleParagraph.ApplyStyle(BuiltinStyle.Title)

paragraph = sec.AddParagraph()
tr = paragraph.AppendText("Strikethough Text")
tr.CharacterFormat.IsStrikeout = True

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Shadow Text")
tr.CharacterFormat.IsShadow = True

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Small caps Text")
tr.CharacterFormat.IsSmallCaps = True

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Double Strikethough Text")
tr.CharacterFormat.DoubleStrike = True

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Outline Text")
tr.CharacterFormat.IsOutLine = True

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("AllCaps Text")
tr.CharacterFormat.AllCaps = True

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Text")
tr = paragraph.AppendText("SubScript")
tr.CharacterFormat.SubSuperScript = SubSuperScript.SubScript

tr = paragraph.AppendText("And")
tr = paragraph.AppendText("SuperScript")
tr.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Emboss Text")
tr.CharacterFormat.Emboss = True
tr.CharacterFormat.TextColor = Color.get_White()

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Hidden:")
tr = paragraph.AppendText("Hidden Text")
tr.CharacterFormat.Hidden = True

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Engrave Text")
tr.CharacterFormat.Engrave = True
tr.CharacterFormat.TextColor = Color.get_White()

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("WesternFonts中文字体")
tr.CharacterFormat.FontNameAscii = "Calibri"
tr.CharacterFormat.FontNameNonFarEast = "Calibri"
tr.CharacterFormat.FontNameFarEast = "Simsun"

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Font Size")
tr.CharacterFormat.FontSize = 20

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Font Color")
tr.CharacterFormat.TextColor = Color.get_Red()

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Bold Italic Text")
tr.CharacterFormat.Bold = True
tr.CharacterFormat.Italic = True

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Underline Style")
tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Highlight Text")
tr.CharacterFormat.HighlightColor = Color.get_Yellow()

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Text has shading")
tr.CharacterFormat.TextBackgroundColor = Color.get_Green()

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Border Around Text")
tr.CharacterFormat.Border.BorderType = BorderStyle.Single

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Text Scale")
tr.CharacterFormat.TextScale = 150

paragraph.AppendBreak(BreakType.LineBreak)
tr = paragraph.AppendText("Character Spacing is 2 point")
tr.CharacterFormat.CharacterSpacing = 2

document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

