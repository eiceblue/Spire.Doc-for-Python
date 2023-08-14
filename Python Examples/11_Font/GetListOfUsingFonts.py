from spire.doc import *
from spire.doc.common import *


def WriteAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s)

class FontInfo:
    def __init__(self):
        self._m_name = ''
        self._m_size = None

    def __eq__(self,other):
        if isinstance(other,FontInfo):
            return self._m_name == other.get_name() and self._m_size == other.get_size()
        return False

    def get_name(self):
        return self._m_name

    def set_name(self, value):
        self._m_name = value

    def get_size(self):
        return self._m_size

    def set_size(self, value):
        self._m_size = value


inputFile = "./Data/UsingFonts.docx"
outputFile = "GetListOfUsingFonts.txt"

fontImformations = ''

font_infos = []

#Load word document
document = Document()
document.LoadFromFile(inputFile)

for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    for j in range(section.Body.Paragraphs.Count):
        paragraph = section.Body.Paragraphs.get_Item(j)
        for k in range(paragraph.ChildObjects.Count):
            obj = paragraph.ChildObjects.get_Item(k)
            if obj.DocumentObjectType is DocumentObjectType.TextRange:
                txtRange = obj if isinstance(obj, TextRange) else None
                tempFont = FontInfo()
                fontName = txtRange.CharacterFormat.FontName
                fontSize = txtRange.CharacterFormat.FontSize
                tempFont.set_name(fontName)
                tempFont.set_size(fontSize)
                if tempFont not in font_infos:
                    font_infos.append(tempFont)
                    textColor = txtRange.CharacterFormat.TextColor.Name
                    s = "Font Name: {0:s}, Size:{1:f}, Color:{2:s}".format(tempFont.get_name(), tempFont.get_size(), textColor)
                    fontImformations += s
                    fontImformations += '\r'

WriteAllText(outputFile, fontImformations)
document.Close()