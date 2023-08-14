from spire.doc import *
from spire.doc.common import *

def InsertHeaderAndFooter(Section):
    header = section.HeadersFooters.Header
    footer = section.HeadersFooters.Footer

    #insert picture and text to header
    headerParagraph = header.AddParagraph()

    headerPicture = headerParagraph.AppendPicture("./Data/Header.png")

    #header text
    text = headerParagraph.AppendText("Demo of Spire.Doc")
    text.CharacterFormat.FontName = "Arial"
    text.CharacterFormat.FontSize = 10
    text.CharacterFormat.Italic = True
    headerParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

    #border
    headerParagraph.Format.Borders.Bottom.BorderType = BorderStyle.Single
    headerParagraph.Format.Borders.Bottom.Space = 0.05


    #header picture layout - text wrapping
    headerPicture.TextWrappingStyle = TextWrappingStyle.Behind

    #header picture layout - position
    headerPicture.HorizontalOrigin = HorizontalOrigin.Page
    headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
    headerPicture.VerticalOrigin = VerticalOrigin.Page
    headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top

    #insert picture to footer
    footerParagraph = footer.AddParagraph()

    footerPicture = footerParagraph.AppendPicture("./Data/Footer.png")

    #footer picture layout
    footerPicture.TextWrappingStyle = TextWrappingStyle.Behind
    footerPicture.HorizontalOrigin = HorizontalOrigin.Page
    footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
    footerPicture.VerticalOrigin = VerticalOrigin.Page
    footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom

    #insert page number
    footerParagraph.AppendField("page number", FieldType.FieldPage)
    footerParagraph.AppendText(" of ")
    footerParagraph.AppendField("number of pages", FieldType.FieldNumPages)
    footerParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

    #border
    footerParagraph.Format.Borders.Top.BorderType = BorderStyle.Single
    footerParagraph.Format.Borders.Top.Space = 0.05

inputFile = "./Data/Sample.docx"
outputFile = "HeaderAndFooter.docx"

# Create word document
document = Document()

document.LoadFromFile(inputFile)
section = document.Sections[0]

# insert header and footer
InsertHeaderAndFooter(section)

# Save as docx file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()



