from spire.doc import *
from spire.doc.common import *

def WriteAllText(fname:str,text:List[str]):
        fp = open(fname,"w")
        for s in text:
            fp.write(s)
        fp.close()


inputFile = "./Data/Template_Docx_3.docx"
outputFile = "GetParagraphByStyleName.txt"
     
#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

content = ""
#content.AppendLine("Get paragraphs by style name \"Heading1\": ")
content += "Get paragraphs by style name \"Heading1\": "
content += "\n"

#Get paragraphs by style name.
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    for j in range(section.Paragraphs.Count):
        paragraph = section.Paragraphs.get_Item(j)
        if paragraph.StyleName == "Heading1":
            #content.AppendLine(paragraph.Text)
            content += paragraph.Text
            content += "\n"

#Save to file.
WriteAllText(outputFile, content)
document.Close()
