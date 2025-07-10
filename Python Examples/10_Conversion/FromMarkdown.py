from spire.doc import *
from spire.doc.common import *

document = Document()
#load .md file
document.LoadFromFile("Data/FromMarkdown.md")
#save to .md file
document.SaveToFile("FromMarkdown_markdown.md", FileFormat.Markdown)
#save to .docx file
document.SaveToFile("FromMarkdown_docx.docx", FileFormat.Docx)
#save to .doc file
document.SaveToFile("FromMarkdown_doc.doc", FileFormat.Doc)
#save to .pdf file
document.SaveToFile("FromMarkdown_pdf.pdf", FileFormat.PDF)
