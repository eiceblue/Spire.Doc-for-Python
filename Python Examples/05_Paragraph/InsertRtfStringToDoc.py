from spire.doc import *
from spire.doc.common import *



outputFile = "InsertRtfStringToDoc.docx"


#Create Word document.
document = Document()

#Add a new section.
section = document.AddSection()

#Add a paragraph to the section.
para = section.AddParagraph()

#Declare a String variable to store the Rtf string.
rtfString = """{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 hakuyoxingshu7000;}}\\f0\\fs28 Hello, World}"""

#Append Rtf string to paragraph.
para.AppendRTF(rtfString)

#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()