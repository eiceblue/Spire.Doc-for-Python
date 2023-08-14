from spire.doc import *
from spire.doc.common import *


def _InsertHyperlink(section):
    paragraph = section.Paragraphs[0] if section.Paragraphs.Count > 0 else section.AddParagraph(
    )
    paragraph.AppendText(
        "Spire.Doc for Python \r\n e-iceblue company Ltd. 2002-2010 All rights reserverd")
    paragraph.ApplyStyle(BuiltinStyle.Heading2)

    paragraph = section.AddParagraph()
    paragraph.AppendText("Home page")
    paragraph.ApplyStyle(BuiltinStyle.Heading2)
    paragraph = section.AddParagraph()
    paragraph.AppendHyperlink(
        "www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink)

    paragraph = section.AddParagraph()
    paragraph.AppendText("Contact US")
    paragraph.ApplyStyle(BuiltinStyle.Heading2)
    paragraph = section.AddParagraph()
    paragraph.AppendHyperlink(
        "mailto:support@e-iceblue.com", "support@e-iceblue.com", HyperlinkType.EMailLink)

    paragraph = section.AddParagraph()
    paragraph.AppendText("Forum")
    paragraph.ApplyStyle(BuiltinStyle.Heading2)
    paragraph = section.AddParagraph()
    paragraph.AppendHyperlink(
        "www.e-iceblue.com/forum/", "www.e-iceblue.com/forum/", HyperlinkType.WebLink)

    paragraph = section.AddParagraph()
    paragraph.AppendText("Download Link")
    paragraph.ApplyStyle(BuiltinStyle.Heading2)
    paragraph = section.AddParagraph()
    paragraph.AppendHyperlink("www.e-iceblue.com/Download/download-word-for-python-now.html",
                              "www.e-iceblue.com/Download/download-word-for-python-now.html", HyperlinkType.WebLink)

    paragraph = section.AddParagraph()
    paragraph.AppendText("Insert Link On Image")
    paragraph.ApplyStyle(BuiltinStyle.Heading2)
    paragraph = section.AddParagraph()
    picture = paragraph.AppendPicture("./Data/Spire.Doc.png")

    paragraph.AppendHyperlink(
        "www.e-iceblue.com/Introduce/doc-for-python.html", picture, HyperlinkType.WebLink)


outputFile = "Hyperlink.docx"

# Open a blank word document as template
document = Document()
section = document.AddSection()

# Insert hyperlink
_InsertHyperlink(section)

# Save doc file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
