from spire.doc import *
from spire.doc.common import *

inputFile1 = "./Data/Spire.Doc.png"
inputFile2 = "./Data/Word.png"

outputFile = "PictureCaptionCrossReference.docx"
#Create word document
document = Document()

#Create a new section
section = document.AddSection()

#Add the first paragraph
firstPara = section.AddParagraph()

#Add the first picture
par1 = section.AddParagraph()
par1.Format.AfterSpacing = 10.0
pic1 = par1.AppendPicture(inputFile1)

pic1.Height = 120.0
pic1.Width = 120.0
#Add caption to the picture
tempFormat = CaptionNumberingFormat.Number
captionParagraph = pic1.AddCaption("Figure", tempFormat, CaptionPosition.BelowItem)
section.AddParagraph()

#Add the second picture
par2 = section.AddParagraph()
pic2 = par2.AppendPicture(inputFile2)

pic2.Height = 120.0
pic2.Width = 120.0
#Add caption to the picture
captionParagraph = pic2.AddCaption("Figure", tempFormat, CaptionPosition.BelowItem)
section.AddParagraph()

#Create a bookmark
bookmarkName = "Figure_2"
paragraph = section.AddParagraph()
paragraph.AppendBookmarkStart(bookmarkName)
paragraph.AppendBookmarkEnd(bookmarkName)

#Replace bookmark content
navigator = BookmarksNavigator(document)
navigator.MoveToBookmark(bookmarkName)
part = navigator.GetBookmarkContent()
part.BodyItems.Clear()
part.BodyItems.Add(captionParagraph)
navigator.ReplaceBookmarkContent(part)

#Create cross-reference field to point to bookmark "Figure_2"
field = Field(document)
field.Type = FieldType.FieldRef
field.Code = """REF Figure_2 \p \h"""
firstPara.ChildObjects.Add(field)
fieldSeparator = FieldMark(document, FieldMarkType.FieldSeparator)
firstPara.ChildObjects.Add(fieldSeparator)

#Set the display text of the field
tr = TextRange(document)
tr.Text = "Figure 2"
firstPara.ChildObjects.Add(tr)

fieldEnd = FieldMark(document, FieldMarkType.FieldEnd)
firstPara.ChildObjects.Add(fieldEnd)

#Update fields
document.IsUpdateFields = True


#Save the Word document
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()