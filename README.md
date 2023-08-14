# Python API for Word Documents Manipulating and Converting

[![Foo](https://i.imgur.com/FxuTGMR.png)](https://www.e-iceblue.com/Introduce/doc-for-python.html)

[Product Page](https://www.e-iceblue.com/Introduce/doc-for-python.html) | Documentation | Examples | [Forum](https://www.e-iceblue.com/forum/spire-doc-f6.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)

[Spire.Doc for Python](https://www.e-iceblue.com/Introduce/doc-for-python.html) is a professional Python API for processing Word documents. Developers can use this independent Python API to create, read, edit, convert, and compare Word documents with any other external dependencies.

[Spire.Doc for Python](https://www.e-iceblue.com/Introduce/doc-for-python.html) is highly compatible with MS Word. It supports Word 97-2003 /2007/2010/2013/2016/2019 and can convert them to commonly used file formats like XML, RTF, TXT, XPS, EPUB, EMF, HTML, and vice versa. Converting Word Doc/Docx to PDF and HTML to image is also supported.

### Standalone Python API
Spire.Doc for Python is an independent Python API to process Word documents. It doesn't requires MS Word or any other third-party tools.

### Support for Old and New Word Versions
- Word 97-03
- Word 2007
- Word 2010
- Word 2013
- Word 2016
- Word 2019

### Rich Support for Word Document Features
A common use of Spire.Doc for Python is to create Word documents dynamically from scratch. Almost all Word document elements are supported, including pages, sections, headers, footers, digital signatures, footnotes, paragraphs, lists, tables, text, fields, hyperlinks, bookmarks, comments, images, style, background settings, document settings, and protection. Furthermore, drawing objects including shapes, text boxes, images, OLE objects, Latex Math Symbols, MathML Code, and controls are supported as well.

### High-Quality File Conversion
By using Spire.Doc for Python, users can save Word Doc/Docx to stream, save as web response and convert Word Doc/Docx to XML, RTF, EMF, TXT, XPS, EPUB, HTML, SVG, ODT, and vice versa. Spire.Doc for Python also supports converting Word Doc/Docx to PDF and HTML to images.

## Examples

### Create a Word document in Python
```Python
﻿from spire.doc.common import *
from spire.doc import *

outputFile = "HelloWorld.docx"
#Create a word document
document = Document()

#Create a new section
section = document.AddSection()

#Create a new paragraph
paragraph = section.AddParagraph()

#Append Text
paragraph.AppendText("Hello World!")

#Save doc file.
document.SaveToFile(outputFile, FileFormat.Docx)

#Close the document object
document.Close()
```

### Convert Word to PDF
```Python
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ConvertedTemplate.docx"
outputFile = "ToPDF.pdf"
        
#Create word document
document = Document()
document.LoadFromFile(inputFile)
#Save the document to a PDF file.
document.SaveToFile(outputFile, FileFormat.PDF)
document.Close()
```

### Convert Word to image
```Python
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ConvertedTemplate.docx"
outputFile =  "ToImage.png"
#Create word document
document = Document()
document.LoadFromFile(inputFile)
#Obtain image data in the default format of png,you can use it to convert other image format.
imageStream = document.SaveImageToStreams(0, ImageType.Bitmap)
with open(outputFile,'wb') as imageFile:
    imageFile.write(imageStream.ToArray())
document.Close()
```

[Product Page](https://www.e-iceblue.com/Introduce/doc-for-python.html) | Documentation | Examples | [Forum](https://www.e-iceblue.com/forum/spire-doc-f6.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)
