from spire.doc import *
from spire.doc.common import *

def IsInNextDocument( element):
    if isinstance(element, Paragraph):
        p = element if isinstance(element, Paragraph) else None
        if p.StyleName == "Heading1":
            return True
    return False

inputFile = "./Data/Template_N5.docx"
outputFolder = "SplitDocIntoHtmlPages/"

#Create Word document.
document = Document()
document.LoadFromFile(inputFile)
subDoc = None
first = True
index = 0
for k in range(document.Sections.Count):
    sec = document.Sections.get_Item(k)
    for m in range(sec.Body.ChildObjects.Count):
        element = sec.Body.ChildObjects.get_Item(m)
        if IsInNextDocument(element):
            if not first:
                #Embed css tyle and image data into html page
                subDoc.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal
                subDoc.HtmlExportOptions.ImageEmbedded = True
                #Save to html file
                result = outputFolder + "out-{0}.html".format(index)
                subDoc.SaveToFile(result, FileFormat.Html)
                index += 1
                subDoc = None
            first = False
        if subDoc is None:
            subDoc = Document()
            subDoc.AddSection()
        subDoc.Sections[0].Body.ChildObjects.Add(element.Clone())
if subDoc is not None:
    #Embed css tyle and image data into html page
    subDoc.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal
    subDoc.HtmlExportOptions.ImageEmbedded = True
    #Save to html file
    resultF = outputFolder+"out-{0}.html".format(index)
    subDoc.SaveToFile(resultF, FileFormat.Html)
    index += 1
    subDoc.Close()
