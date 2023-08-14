from spire.doc import *
from spire.doc.common import *
import os
import sys

def WriteAllText(fname:str,text:List[str]):
        fp = open(fname,"w")
        for s in text:
            fp.write(s)
        fp.close()


inputFile = "./Data/ExtractParagraphBasedOnStyle.docx"
outputFile = "ExtractParagraphBasedOnStyle.txt"
     

#Create a new document
document = Document()
styleName1 = "Heading1"
style1Text = ''
#Load file from disk
document.LoadFromFile(inputFile)

style1Text += "The following is the content of the paragraph with the style name "
style1Text += styleName1
style1Text += ": "
style1Text += "\n"
#Extrct paragraph based on style
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    for j in range(section.Paragraphs.Count):
        paragraph = section.Paragraphs.get_Item(j)
        if paragraph.StyleName is not None and paragraph.StyleName == styleName1:
             style1Text += paragraph.Text
             style1Text += '\n'
WriteAllText(outputFile, style1Text)
document.Close()     