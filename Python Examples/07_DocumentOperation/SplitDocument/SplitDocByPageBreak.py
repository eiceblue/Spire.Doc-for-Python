from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/SplitWordFileByPageBreak.docx"
outputFolder = "SplitDocByPageBreak/"

#Create Word document.
original = Document()
#Load the file from disk.
original.LoadFromFile(inputFile)
#Create a new word document and add a section to it.
newWord =Document()
section = newWord.AddSection()
original.CloneDefaultStyleTo(newWord)
original.CloneThemesTo(newWord)
original.CloneCompatibilityTo(newWord)
#Split the original word document into separate documents according to page break.
index = 0
#Traverse through all sections of original document.
for m in range(original.Sections.Count):
    sec = original.Sections.get_Item(m)
    #Traverse through all body child objects of each section.
    for k in range(sec.Body.ChildObjects.Count):
        obj = sec.Body.ChildObjects.get_Item(k)
        if isinstance(obj, Paragraph):
            para = obj if isinstance(obj, Paragraph) else None
            sec.CloneSectionPropertiesTo(section)
            #Add paragraph object in original section into section of new document.
            section.Body.ChildObjects.Add(para.Clone())
            for j in range(para.ChildObjects.Count):
                parobj = para.ChildObjects.get_Item(j)
                if isinstance(parobj, Break) and ( parobj if isinstance(parobj, Break) else None).BreakType == BreakType.PageBreak:
                    #Get the index of page break in paragraph.
                    i = para.ChildObjects.IndexOf(parobj)
                    #Remove the page break from its paragraph.
                    section.Body.LastParagraph.ChildObjects.RemoveAt(i)
                    #Save the new document to a Docx file.
                    resultF = outputFolder
                    resultF += "SplitDocByPageBreak-{0}.docx".format(index)
                    newWord.SaveToFile(resultF, FileFormat.Docx)
                    index += 1
                    #Create a new document and add a section.
                    newWord = Document()
                    section = newWord.AddSection()
                    original.CloneDefaultStyleTo(newWord)
                    original.CloneThemesTo(newWord)
                    original.CloneCompatibilityTo(newWord)
                    sec.CloneSectionPropertiesTo(section)
                    #Add paragraph object in original section into section of new document.
                    section.Body.ChildObjects.Add(para.Clone())
                    if section.Paragraphs[0].ChildObjects.Count == 0:
                        #Remove the first blank paragraph.
                        section.Body.ChildObjects.RemoveAt(0)
                    else:
                        #Remove the child objects before the page break.
                        while i >= 0:
                            section.Paragraphs[0].ChildObjects.RemoveAt(i)
                            i -= 1
        if isinstance(obj, Table):
            #Add table object in original section into section of new document.
            section.Body.ChildObjects.Add(obj.Clone())            
#Save the file.
result = outputFolder+"SplitDocByPageBreak-{0}.docx".format(index)
newWord.SaveToFile(result, FileFormat.Docx2013)
newWord.Close()