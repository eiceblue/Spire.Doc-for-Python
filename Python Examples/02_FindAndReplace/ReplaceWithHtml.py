from spire.doc import *
from spire.doc.common import *

class TextRangeLocation:
    def __init__(self, text):
         self._m_Text = None
         self.set_text(text)

    def get_text(self):
         return self._m_Text
    def set_text(self, value):
         self._m_Text = value

    def get_owner(self):
        return self.get_text().OwnerParagraph

    def get_index(self):
        return self.get_owner().ChildObjects.IndexOf(self.get_text())

    def CompareTo(self, other):
        return -(self.get_index() - other.get_index())


def ReplaceWithHTML(location, replacement):
    textRange = location.get_text()

    #textRange index
    index = location.get_index()

    #get owener paragraph
    paragraph = location.get_owner()

    #get owner text body
    sectionBody = paragraph.OwnerTextBody

    #get the index of paragraph in section
    paragraphIndex = sectionBody.ChildObjects.IndexOf(paragraph)

    replacementIndex = -1
    if index == 0:
        #remove the first child object
        paragraph.ChildObjects.RemoveAt(0)

        replacementIndex = sectionBody.ChildObjects.IndexOf(paragraph)
    elif index == paragraph.ChildObjects.Count - 1:
        paragraph.ChildObjects.RemoveAt(index)
        replacementIndex = paragraphIndex + 1
    else:
        #split owner paragraph
        paragraph1 = paragraph.Clone()
        while paragraph.ChildObjects.Count > index:
            paragraph.ChildObjects.RemoveAt(index)
            i = 0
        count = index + 1
        while i < count:
            paragraph1.ChildObjects.RemoveAt(0)
            i += 1
        sectionBody.ChildObjects.Insert(paragraphIndex + 1, paragraph1)

        replacementIndex = paragraphIndex + 1

    #insert replacement
    i = 0
    while i <= len(replacement) - 1:
        sectionBody.ChildObjects.Insert(replacementIndex + i, replacement[i].Clone())
        i += 1

inputFile = "./Data/ReplaceWithHtml.docx"
inputHTMLFile = "./Data/InputHtml1.txt"
outputFile = "ReplaceWithHtml.docx"

#Load the document from disk.  
document = Document()
document.LoadFromFile(inputFile)

#collect the objects which is used to replace text
replacement = []

#create a temporary section
tempSection = document.AddSection()

with open(inputHTMLFile, 'r') as f:
    HTML = f.read()
#add a paragraph to append html
par = tempSection.AddParagraph()
par.AppendHTML(HTML)

#get the objects in temporary section
for i in range(tempSection.Body.ChildObjects.Count):
    obj = tempSection.Body.ChildObjects.get_Item(i)
    docObj = obj if isinstance(obj, DocumentObject) else None
    replacement.append(docObj)

#Find all text which will be replaced.
selections = document.FindAllString("[#placeholder]", False, True)

locations = []
for selection in selections:
    locations.append(TextRangeLocation(selection.GetAsOneRange()))
locations.sort()

for location in locations:
    #replace the text with HTML
    ReplaceWithHTML(location, replacement)

#remove the temp section
document.Sections.Remove(tempSection)

#Save the document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()

