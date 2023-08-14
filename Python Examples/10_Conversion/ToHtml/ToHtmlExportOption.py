from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/ToHtmlTemplate.docx"
outputFile = "ToHtmlExportOption.html"
#Open a Word document.
document = Document()
document.LoadFromFile(inputFile)
#Set whether the css styles are embeded or not. 
document.HtmlExportOptions.CssStyleSheetFileName = "sample.css"
document.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.External
#Set whether the images are embeded or not. 
document.HtmlExportOptions.ImageEmbedded = False
document.HtmlExportOptions.ImagesPath = "./"
#Set the option whether to export form fields as plain text or not.
document.HtmlExportOptions.IsTextInputFormFieldAsText = True
#Save the document to a html file.
document.SaveToFile(outputFile, FileFormat.Html)
document.Close()
