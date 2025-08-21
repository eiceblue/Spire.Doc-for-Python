# Python代码核心功能提取

# Spire.Doc Python Hello World
## Create a simple Word document with Hello World text
```python
#Create a word document
document = Document()

#Create a new section
section = document.AddSection()

#Create a new paragraph
paragraph = section.AddParagraph()

#Append Text
paragraph.AppendText("Hello World!")
```

---

# spire.doc python find and highlight
## Find and highlight specific text in a Word document
```python
#Find text
textSelections = document.FindAllString("word", False, True)

#Set highlight
for selection in textSelections:
    selection.GetAsOneRange().CharacterFormat.HighlightColor = Color.get_Yellow()
```

---

# Spire.Doc Get Ranges Functionality
## Demonstrates how to get ranges from text selections and apply formatting
```python
#Find text
textSelections = document.FindAllString("Spire.Doc", False, True)

#test GetRanges()
textRanges = textSelections[0].GetRanges()
textRanges[0].CharacterFormat.HighlightColor = Color.get_Yellow()

#test GetAsRange()
textRange= textSelections[1].GetAsRange()
textRange[0].CharacterFormat.HighlightColor = Color.get_Red()

#test GetAsRange(bool IsCopyFormat)
textRange=textSelections[2].GetAsRange(True)
textRange[0].CharacterFormat.HighlightColor = Color.get_Green()
```

---

# Spire.Doc Python Document Content Replacement
## Replace specific text in a document with content from another document
```python
#Get the first section of the first document 
section1 = document1.Sections[0]

#Create a regex
regex = Regex("""\\[MY_DOCUMENT\\]""", RegexOptions.none)

#Find the text by regex
textSections = document1.FindAllPattern(regex)

#Travel the found strings
for seletion in textSections:

    #Get the paragraph
    para = seletion.GetAsOneRange().OwnerParagraph

    #Get textRange
    textRange = seletion.GetAsOneRange()

    #Get the paragraph index
    index = section1.Body.ChildObjects.IndexOf(para)

    #Insert the paragraphs of document2
    for i in range(document2.Sections.Count):
        section2 = document2.Sections.get_Item(i)
        for j in range(section2.Paragraphs.Count):
            paragraph = section2.Paragraphs.get_Item(j)
            section1.Body.ChildObjects.Insert(index, paragraph.Clone() if isinstance(paragraph.Clone(), Paragraph) else None)
            #Remove the found textRange
            para.ChildObjects.Remove(textRange)
```

---

# Spire.Doc Python Regex Text Replacement
## Replace text in document using regular expressions
```python
#create a regex, match the text that starts with #
regex = Regex("""\\#\\w+\\b""")

#replace the text by regex
document.Replace(regex, "Spire.Doc")
```

---

# Spire.Doc Python Text Replacement
## Replace text with field in Word document
```python
#Find the target text
selection = document.FindString('summary', False, True)

#Get text range
textRange = selection.GetAsOneRange()
#Get it's owner paragraph
ownParagraph:Paragraph = textRange.OwnerParagraph
#Get the index of this text range
rangeIndex = ownParagraph.ChildObjects.IndexOf(textRange)
#Remove the text range
ownParagraph.ChildObjects.RemoveAt(rangeIndex)
#Remove the objects which are behind the text range
tempList = []
for i in range(rangeIndex, ownParagraph.ChildObjects.Count):
    #Add a copy of these objects into a temp list
    tempP = ownParagraph.ChildObjects.get_Item(rangeIndex).Clone()
    tempList.append(tempP)
    ownParagraph.ChildObjects.RemoveAt(rangeIndex)
#Append field to the paragraph
ownParagraph.AppendField("MyFieldName", FieldType.FieldMergeField)
#Put these objects back into the paragraph one by one
for obj in tempList:
    ownParagraph.ChildObjects.Add(obj)
```

---

# spire.doc python replace text with table
## Replace specific text with a table in a Word document
```python
#Create Word document.
document=Document()

#Return TextSection by finding the key text string "ChristmasDay,December25".
section=document.Sections[0]
selection=document.FindString("Christmas Day, December 25",True,True)

#Return TextRange from TextSection, then get OwnerParagraph through TextRange.
range=selection.GetAsOneRange()
paragraph=range.OwnerParagraph

#Return the zero-based index of the specified paragraph.
body=paragraph.OwnerTextBody
index=body.ChildObjects.IndexOf(paragraph)

#Create a new table.
table=section.AddTable(True)
table.ResetCells(3,3)

#Remove the paragraph and insert table into the collection at the specified index.
body.ChildObjects.Remove(paragraph)
body.ChildObjects.Insert(index,table)
```

---

# spire.doc python replace text with document
## Replace specified text with another document
```python
#Load a template document 
doc = Document("template.docx")

#Load another document to replace text
replaceDoc = Document("replacement.docx")
#Replace specified text with the other document
doc.Replace("Document1", replaceDoc, False, True)
```

---

# spire.doc python replace text with html
## TextRangeLocation class and ReplaceWithHTML function for replacing text with HTML content
```python
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

    #get owner paragraph
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
```

---

# Spire.Doc Python Text Replacement
## Replace text with image in document
```python
#Find the string "E-iceblue" in the document
selections = doc.FindAllString("E-iceblue", True, True)
index = 0
testRange = None

#Remove the text and replace it with Image
for selection in selections:
    pic = DocPicture(doc)
    pic.LoadImage(inputFile2)

    testRange = selection.GetAsOneRange()
    index = testRange.OwnerParagraph.ChildObjects.IndexOf(testRange)
    testRange.OwnerParagraph.ChildObjects.Insert(index, pic)
    testRange.OwnerParagraph.ChildObjects.Remove(testRange)
```

---

# spire.doc python text replacement
## Replace specific text in a Word document
```python
#Create word document
document = Document()

#Replace text
document.Replace("word", "ReplacedText", False, True)
```

---

# spire.doc extract content between paragraphs
## Extract content between specified paragraphs from one document to another
```python
def ExtractBetweenParagraphs(sourceDocument, destinationDocument, startPara, endPara):
    #Extract the content
    for i in range(startPara - 1, endPara):
        #Clone the ChildObjects of source document
        doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone()

        #Add to destination document 
        destinationDocument.Sections[0].Body.ChildObjects.Add(doobj)
```

---

# Spire.Doc Python Content Extraction
## Extract content between paragraphs with specific styles
```python
def ExtractBetweenParagraphStyles(sourceDocument, destinationDocument, stylename1, stylename2):
        startindex = 0
        endindex = 0
        #travel the sections of source document
        for i in range(sourceDocument.Sections.Count):
            section = sourceDocument.Sections.get_Item(i)
            #travel the paragraphs
            for j in range(section.Paragraphs.Count):
                paragraph = section.Paragraphs.get_Item(j)
                #Judge paragraph style1
                if paragraph.StyleName == stylename1:
                    #Get the paragraph index
                    startindex = section.Body.Paragraphs.IndexOf(paragraph)
                #Judge paragraph style2
                if paragraph.StyleName == stylename2:
                    #Get the paragraph index
                    endindex = section.Body.Paragraphs.IndexOf(paragraph)
            #Extract the content
            for k in range(startindex + 1, endindex):
                #Clone the ChildObjects of source document
                doobj = sourceDocument.Sections[0].Body.ChildObjects[k].Clone()

                #Add to destination document 
                destinationDocument.Sections[0].Body.ChildObjects.Add(doobj)
```

---

# spire.doc python extract paragraphs by style
## extract paragraphs from document based on style name
```python
#Create a new document
document = Document()
styleName1 = "Heading1"
style1Text = ''

#Extract paragraph based on style
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    for j in range(section.Paragraphs.Count):
        paragraph = section.Paragraphs.get_Item(j)
        if paragraph.StyleName is not None and paragraph.StyleName == styleName1:
             style1Text += paragraph.Text
             style1Text += '\n'
document.Close()
```

---

# Extract content from bookmark in Word document
## This code demonstrates how to extract content from a bookmark in a Word document using Spire.Doc for Python

```python
# Create bookmark navigator and locate bookmark
navigator = BookmarksNavigator(sourcedocument)
navigator.MoveToBookmark("Test", True, True)

# Get bookmark content
textBodyPart = navigator.GetBookmarkContent()

# Create a TextRange type list
listTextRange = []

# Traverse the items of text body
for i in range(textBodyPart.BodyItems.Count):
    item = textBodyPart.BodyItems.get_Item(i)
    # if it is paragraph
    if isinstance(item, Paragraph):
        tempItems = item.ChildObjects
        # Traverse the ChildObjects of the paragraph
        for j in range(tempItems.Count):
            childObject = tempItems.get_Item(j)
            # if it is TextRange
            if isinstance(childObject, TextRange):
                # Add it into list
                textRange = childObject
                listTextRange.append(textRange)

# Add the extract content to destination document
for m, unusedItem in enumerate(listTextRange):
    paragraph.Items.Add(listTextRange[m].Clone())
```

---

# Extract Content from Comment Range
## This code demonstrates how to extract content from a comment range in a document using Spire.Doc for Python.
```python
#Create a document
sourceDoc = Document()

#Create a destination document
destinationDoc = Document()

#Add section for destination document
destinationSec = destinationDoc.AddSection()

#Get the first comment
comment = sourceDoc.Comments[0]

#Get the paragraph of obtained comment
para = comment.OwnerParagraph

#Get index of the CommentMarkStart 
startIndex = para.ChildObjects.IndexOf(comment.CommentMarkStart)

#Get index of the CommentMarkEnd
endIndex = para.ChildObjects.IndexOf(comment.CommentMarkEnd)

paragraph = destinationSec.AddParagraph()

#Traverse paragraph ChildObjects
for i in range(startIndex, endIndex + 1):
    #Clone the ChildObjects of source document
    dobj = para.ChildObjects[i].Clone()

    #Add to destination document 
    paragraph.ChildObjects.Add(dobj)
```

---

# spire.doc python content extraction
## extract content from paragraph to table
```python
def ExtractByTable(sourceDocument, destinationDocument, startPara, tableNo):
    #Get the table from the source document
    table = sourceDocument.Sections[0].Tables[tableNo - 1] if isinstance(sourceDocument.Sections[0].Tables[tableNo - 1], Table) else None

    #Get the table index
    index = sourceDocument.Sections[0].Body.ChildObjects.IndexOf(table)
    for i in range(startPara - 1, index + 1):
        #Clone the ChildObjects of source document
        doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone()

        #Add to destination document 
        destinationDocument.Sections[0].Body.ChildObjects.Add(doobj)
```

---

# spire.doc python extract content
## extract content starting from form field
```python
#Create the source document
sourceDocument = Document()

#Create a destination document
destinationDoc = Document()

#Add a section
section = destinationDoc.AddSection()

#Define a variables
index = 0

#Traverse FormFields
for i in range(sourceDocument.Sections[0].Body.FormFields.Count):
    #Find FieldFormTextInput type field
    field = sourceDocument.Sections[0].Body.FormFields.get_Item(i)
    if field.Type == FieldType.FieldFormTextInput:
        #Get the paragraph
        paragraph = field.OwnerParagraph

        #Get the index
        index = sourceDocument.Sections[0].Body.ChildObjects.IndexOf(paragraph)
        break

#Extract the content
i = index
while i < index + 3:
    #Clone the ChildObjects of source document
    doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone()

    #Add to destination document 
    section.Body.ChildObjects.Add(doobj)
    i += 1
```

---

# spire.doc document sections management
## add and delete sections in a document
```python
#Add a section
doc.AddSection()

#Delete the last section
doc.Sections.RemoveAt(doc.Sections.Count - 1)
```

---

# spire.doc python section clone
## clone section from one document to another
```python
cloneSection = None
for i in range(srcDoc.Sections.Count):
    section = srcDoc.Sections.get_Item(i)
    #Clone section
    cloneSection = section.Clone()
    #Add the cloneSection in destination file
    desDoc.Sections.Add(cloneSection)
```

---

# Spire.Doc Python Section Cloning
## Clone content from one section to another in a Word document
```python
#Get the first section
sec1 = doc.Sections[0]
#Get the second section
sec2 = doc.Sections[1]

#Loop through the contents of sec1
for i in range(sec1.Body.ChildObjects.Count):
    obj = sec1.Body.ChildObjects.get_Item(i)
    #Clone the contents to sec2
    sec2.Body.ChildObjects.Add(obj.Clone())
```

---

# Spire.Doc Python Section Page Setup
## Modify page setup properties of document sections
```python
#Loop through all sections
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    #Modify the margins
    section.PageSetup.Margins = MarginsF(100.0, 80.0, 100.0, 80.0)
    #Modify the page size
    section.PageSetup.PageSize = PageSize.Letter()
```

---

# spire.doc python section manipulation
## remove section content from Word document
```python
#Loop through all sections
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    #Remove header content
    section.HeadersFooters.Header.ChildObjects.Clear()
    #Remove body content
    section.Body.ChildObjects.Clear()
    #Remove footer content
    section.HeadersFooters.Footer.ChildObjects.Clear()
```

---

# Spire.Doc Python Paragraph Tab Stops
## Add tab stops to paragraphs with different alignments and leader types
```python
# Add paragraph
paragraph = section.AddParagraph()

# Add first tab with position and alignment
tab = paragraph.Format.Tabs.AddTab(28.0)
tab.Justification = TabJustification.Left
paragraph.AppendText("\tWashing Machine")

# Add second tab with position, alignment and leader type
tab = paragraph.Format.Tabs.AddTab(280.0)
tab.Justification = TabJustification.Left
tab.TabLeader = TabLeader.Dotted
paragraph.AppendText("\t$650")

# Add another paragraph with different tab leader
paragraph2 = section.AddParagraph()

# Add first tab with position and alignment
tab = paragraph2.Format.Tabs.AddTab(28.0)
tab.Justification = TabJustification.Left
paragraph2.AppendText("\tRefrigerator")

# Add second tab with position, alignment and no leader
tab = paragraph2.Format.Tabs.AddTab(280.0)
tab.Justification = TabJustification.Left
tab.TabLeader = TabLeader.NoLeader
paragraph2.AppendText("\t$800")
```

---

# spire.doc python text wrap
## Allow Latin text to wrap in the middle of a word
```python
para = document.Sections[0].Paragraphs[0]
#Allow Latin text to wrap in the middle of a word
para.Format.WordWrap = False
```

---

# Spire.Doc Paragraph Copying
## Copy paragraphs between Word documents and add a watermark
```python
#Create Word document1.
document1 = Document()

#Create a new document.
document2 = Document()

#Get paragraph 1 and paragraph 2 in document1.
s = document1.Sections[0]
p1 = s.Paragraphs[0]
p2 = s.Paragraphs[1]

#Copy p1 and p2 to document2.
s2 = document2.AddSection()
NewPara1 = p1.Clone()
s2.Paragraphs.Add(NewPara1)
NewPara2 = p2.Clone()
s2.Paragraphs.Add(NewPara2)

#Add watermark.
WM = PictureWatermark()
WM.SetPicture("./Data/Logo.jpg")
document2.Watermark = WM
```

---

# Spire.Doc Python Catalog Creation
## Create a catalog with different heading levels and list styles
```python
#Create Word document.
document = Document()

#Add a new section. 
section = document.AddSection()
paragraph = section.Paragraphs[0] if section.Paragraphs.Count > 0 else section.AddParagraph()

#Add Heading 1.
paragraph = section.AddParagraph()
paragraph.AppendText(BuiltinStyle.Heading1.name)
paragraph.ApplyStyle(BuiltinStyle.Heading1)
paragraph.ListFormat.ApplyNumberedStyle()

#Add Heading 2.
paragraph = section.AddParagraph()
paragraph.AppendText(BuiltinStyle.Heading2.name)
paragraph.ApplyStyle(BuiltinStyle.Heading2)

#List style for Headings 2.
listSty2 = ListStyle(document, ListType.Numbered)
for i in range(listSty2.Levels.Count):
    listLev = listSty2.Levels.get_Item(i)
    listLev.UsePrevLevelPattern = True
    listLev.NumberPrefix = "1."
listSty2.Name = "MyStyle2"
document.ListStyles.Add(listSty2)
paragraph.ListFormat.ApplyStyle(listSty2.Name)

#Add list style 3.
listSty3 = ListStyle(document, ListType.Numbered)
for i in range(listSty3.Levels.Count):
    listLev = listSty3.Levels.get_Item(i)
    listLev.UsePrevLevelPattern = True
    listLev.NumberPrefix = "1.1."
listSty3.Name = "MyStyle3"
document.ListStyles.Add(listSty3)

#Add Heading 3.
for i in range(0, 4):
    paragraph = section.AddParagraph()

    #Append text
    paragraph.AppendText(BuiltinStyle.Heading3.name)

    #Apply list style 3 for Heading 3
    paragraph.ApplyStyle(BuiltinStyle.Heading3)
    paragraph.ListFormat.ApplyStyle(listSty3.Name)
```

---

# Spire.Doc Python Paragraph
## Get paragraphs by style name and extract text
```python
#Get paragraphs by style name.
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    for j in range(section.Paragraphs.Count):
        paragraph = section.Paragraphs.get_Item(j)
        if paragraph.StyleName == "Heading1":
            #Extract text from paragraph with Heading1 style
            text = paragraph.Text
```

---

# Spire.Doc Python Paragraph Revisions
## Extract details of paragraph and text range revisions from a Word document
```python
# Iterate through sections and paragraphs to check for revisions
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    for j in range(section.Paragraphs.Count):
        paragraph = section.Paragraphs.get_Item(j)
        
        # Check if paragraph has delete revision
        if paragraph.IsDeleteRevision:
            builder += "The section {} paragraph {} has been changed (deleted).".format(document.GetIndex(section), section.GetIndex(paragraph))
            builder += "\nAuthor: " + paragraph.DeleteRevision.Author
            builder += "\nDateTime: " + paragraph.DeleteRevision.DateTime.ToString()
            builder += "\nType: " + paragraph.DeleteRevision.Type.name + "\n"
            
        # Check if paragraph has insert revision
        elif paragraph.IsInsertRevision:
            builder += "The section {} paragraph {} has been changed (inserted).".format(document.GetIndex(section), section.GetIndex(paragraph))
            builder += "\nAuthor: " + paragraph.InsertRevision.Author
            builder += "\nDateTime: " + paragraph.InsertRevision.DateTime.ToString()
            builder += "\nType: " + paragraph.InsertRevision.Type.name + "\n"
            
        # Check text ranges within the paragraph for revisions
        else:
            for k in range(paragraph.ChildObjects.Count):
                obj = paragraph.ChildObjects.get_Item(k)
                if obj.DocumentObjectType is DocumentObjectType.TextRange:
                    textRange = obj if isinstance(obj, TextRange) else None
                    
                    # Check if text range has delete revision
                    if textRange.IsDeleteRevision:
                        builder += "The section {} paragraph {} textrange {} has been changed (deleted).".format(document.GetIndex(section), section.GetIndex(paragraph), paragraph.GetIndex(textRange))
                        builder += "\nAuthor: " + textRange.DeleteRevision.Author
                        builder += "\nDateTime: " + textRange.DeleteRevision.DateTime.ToString()
                        builder += "\nType: " + textRange.DeleteRevision.Type.name
                        builder += "\nChange Text: " + textRange.Text + "\n"
                        
                    # Check if text range has insert revision
                    elif textRange.IsInsertRevision:
                        builder += "The section {} paragraph {} textrange {} has been changed (inserted).".format(document.GetIndex(section), section.GetIndex(paragraph), paragraph.GetIndex(textRange))
                        builder += "\nAuthor: " + textRange.InsertRevision.Author
                        builder += "\nDateTime: " + textRange.InsertRevision.DateTime.ToString()
                        builder += "\nType: " + textRange.InsertRevision.Type.name
                        builder += "\nChange Text: " + textRange.Text + "\n"
```

---

# Spire.Doc Python Paragraph Manipulation
## Hide paragraph text in a Word document
```python
# Get the first section and the first paragraph from the word document
sec = document.Sections[0]
para = sec.Paragraphs[0]

# Loop through the text ranges and set CharacterFormat.Hidden property as true to hide the texts
for i in range(para.ChildObjects.Count):
    obj = para.ChildObjects.get_Item(i)
    if isinstance(obj, TextRange):
        trange = obj if isinstance(obj, TextRange) else None
        trange.CharacterFormat.Hidden = True
```

---

# Spire.Doc Python RTF Insertion
## Insert RTF string into Word document
```python
# Create Word document
document = Document()

# Add a new section
section = document.AddSection()

# Add a paragraph to the section
para = section.AddParagraph()

# Declare a String variable to store the Rtf string
rtfString = """{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 hakuyoxingshu7000;}}\\f0\\fs28 Hello, World}"""

# Append Rtf string to paragraph
para.AppendRTF(rtfString)
```

---

# Spire.Doc Python Paragraph Pagination
## Manage paragraph pagination by setting page break before
```python
#Get the first section and the paragraph we want to manage the pagination.
sec = document.Sections[0]
para = sec.Paragraphs[4]

#Set the pagination format as Format.PageBreakBefore for the checked paragraph.
para.Format.PageBreakBefore = True
```

---

# Spire.Doc Python Remove Paragraphs
## This code demonstrates how to remove all paragraphs from a Word document
```python
#Create Word document
document = Document()

#Remove paragraphs from the body of every section in the document
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    section.Paragraphs.Clear()
```

---

# spire.doc python remove empty lines
## remove empty paragraphs from Word document
```python
# Traverse every section on the word document and remove the null and empty paragraphs
for k in range(document.Sections.Count):
    section = document.Sections.get_Item(k)
    i = 0
    while i < section.Body.ChildObjects.Count:
        if section.Body.ChildObjects[i].DocumentObjectType == DocumentObjectType.Paragraph:
            objItem = section.Body.ChildObjects[i]
            if isinstance(objItem, Paragraph):
                paraObj = Paragraph(objItem)
                if len(paraObj.Text) == 0:
                    section.Body.ChildObjects.Remove(objItem)
                    i -= 1
        i += 1
```

---

# Spire.Doc Python paragraph removal
## Remove a specific paragraph from a Word document
```python
#Remove the first paragraph from the first section of the document.
document.Sections[0].Paragraphs.RemoveAt(0)
```

---

# spire.doc paragraph formatting
## Set first line indentation for a paragraph
```python
# Create a Paragraph object using the document
para = Paragraph(document)

# Set the first line indent to 0 characters
para.Format.SetFirstLineIndentChars(0)

# Insert the paragraph at index 1 in the first section of the document
document.Sections[0].Paragraphs.Insert(1, para)
```

---

# spire.doc python frame positioning
## set frame position in document
```python
#Get a paragraph
paragraph = document.Sections[0].Paragraphs[0]

#Set the Frame's position
if paragraph.IsFrame:
    paragraph.Frame.SetHorizontalPosition(150)
    paragraph.Frame.SetVerticalPosition(150)
```

---

# Set Paragraph Shading in Spire.Doc for Python
## Set background color for paragraphs and text in a Word document
```python
#Get a paragraph.
paragaph = document.Sections[0].Paragraphs[0]

#Set background color for the paragraph.
paragaph.Format.BackColor = Color.get_Yellow()

#Set background color for the selected text of paragraph.
paragaph = document.Sections[0].Paragraphs[2]
selection = paragaph.Find("Christmas", True, False)
trange = selection.GetAsOneRange()
trange.CharacterFormat.TextBackgroundColor = Color.get_Yellow()
```

---

# Spire.Doc for Python - Set Snap To Grid
## Demonstrates how to set the SnapToGrid property of a paragraph in a Word document
```python
# Create a new instance of the Document class.
doc = Document()

# Add a new section to the document.
section = doc.AddSection()

# Set the grid type of the page setup in the section to "LinesOnly".
section.PageSetup.GridType = GridPitchType.LinesOnly

# Set the number of lines per page in the section to 15.
section.PageSetup.LinesPerPage = 15

# Add a new paragraph to the section.
paragraph = section.AddParagraph()

# Set the "SnapToGrid" property of the paragraph's format to true.
paragraph.Format.SnapToGrid = True
```

---

# Spire.Doc Python Text Spacing
## Set space between Asian and Latin text
```python
# Get the first paragraph from the first section
para = document.Sections[0].Paragraphs[0]

# Set whether to automatically adjust space between Asian text and Latin text
para.Format.AutoSpaceDE = False
# Set whether to automatically adjust space between Asian text and numbers
para.Format.AutoSpaceDN = True
```

---

# Spire.Doc Python Paragraph Spacing
## Set paragraph spacing before and after in Word document
```python
#Create Word document.
document = Document()

#Add the text strings to the paragraph and set the style.
para = Paragraph(document)
textRange1 = para.AppendText("This is an inserted paragraph.")
textRange1.CharacterFormat.TextColor = Color.get_Blue()
textRange1.CharacterFormat.FontSize = 15

#set the spacing before and after.
para.Format.BeforeAutoSpacing = False
para.Format.BeforeSpacing = 10
para.Format.AfterAutoSpacing = False
para.Format.AfterSpacing = 10

#insert the added paragraph to the word document.
document.Sections[0].Paragraphs.Insert(1, para)
```

---

# Apply Emphasis Mark in Word Document
## This code demonstrates how to find specific text in a Word document and apply emphasis mark to it.
```python
#Find text to emphasize
textSelections = doc.FindAllString("Spire.Doc for Python", False, True)

#Set emphasis mark to the found text
for selection in textSelections:
    selection.GetAsOneRange().CharacterFormat.EmphasisMark = Emphasis.Dot
```

---

# Spire.Doc Python Text Case Change
## Change text case to AllCaps and SmallCaps
```python
# Get the first paragraph and set its CharacterFormat to AllCaps
para1 = doc.Sections[0].Paragraphs[1]

for i in range(para1.ChildObjects.Count):
    obj = para1.ChildObjects.get_Item(i)
    if isinstance(obj, TextRange):
        textRange = obj if isinstance(obj, TextRange) else None
        textRange.CharacterFormat.AllCaps = True

# Get the third paragraph and set its CharacterFormat to IsSmallCaps
para2 = doc.Sections[0].Paragraphs[3]
for i in range(para2.ChildObjects.Count):
    obj = para1.ChildObjects.get_Item(i)
    if isinstance(obj, TextRange):
        textRange = obj if isinstance(obj, TextRange) else None
        textRange.CharacterFormat.IsSmallCaps = True
```

---

# Spire.Doc Python Barcode Creation
## Create a barcode in a Word document
```python
#Create a document
doc = Document()

#Add a paragraph
p = doc.AddSection().AddParagraph()

#Add barcode and set its format
txtRang = p.AppendText("H63TWX11072")
#Set barcode font name, note you need to install the barcode font on your system at first
txtRang.CharacterFormat.FontName = "C39HrP60DlTt"
txtRang.CharacterFormat.FontSize = 80
txtRang.CharacterFormat.TextColor = Color.get_SeaGreen()
```

---

# spire.doc python text extraction
## extract text content from a word document
```python
#Load the document
document = Document()
document.LoadFromFile(inputFile)

#get text from document
text = document.GetText()

#close document
document.Close()
```

---

# spire.doc python text manipulation
## insert and highlight new text in a document
```python
#Find all the text string "Word" from the sample document
selections = doc.FindAllString("Word", True, True)
index = 0
#Defines text range
trange = TextRange(doc)
#Insert new text string (New) after the searched text string
for selection in selections:
    trange = selection.GetAsOneRange()
    newrange = TextRange(doc)
    newrange.Text = ("(New text)")
    index = trange.OwnerParagraph.ChildObjects.IndexOf(trange)
    trange.OwnerParagraph.ChildObjects.Insert(index + 1, newrange)
#Find and highlight the newly added text string New
text = doc.FindAllString("New text", True, True)
for seletion in text:
    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.get_Yellow()
```

---

# spire.doc python insert symbol
## insert unicode symbols into Word document
```python
#Create Word document.
document = Document()
#Add a section.
section = document.AddSection()
#Add a paragraph.
paragraph = section.AddParagraph()
#Use unicode characters to create symbol Ä.
tr = paragraph.AppendText(str('\u00c4'))
#Set the color of symbol Ä.
tr.CharacterFormat.TextColor = Color.get_Red()
#Add symbol Ë.
paragraph.AppendText(str('\u00cb'))
```

---

# Spire.Doc for Python
## Load text with specific encoding
```python
#Create word document
document = Document()
#Load the text file 
document.LoadText(inputFile, Encoding.get_UTF7())
```

---

# spire.doc python superscript subscript
## Set superscript and subscript text in a document
```python
# Add text to paragraph
paragraph.AppendText("E = mc")
range1 = paragraph.AppendText("2")
# Set superscript
range1.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript
paragraph.AppendBreak(BreakType.LineBreak)
paragraph.AppendText("F")
range2 = paragraph.AppendText("n")
# Set subscript
range2.CharacterFormat.SubSuperScript = SubSuperScript.SubScript
paragraph.AppendText(" = F")
paragraph.AppendText("n-1").CharacterFormat.SubSuperScript = SubSuperScript.SubScript
paragraph.AppendText(" + F")
paragraph.AppendText("n-2").CharacterFormat.SubSuperScript = SubSuperScript.SubScript
```

---

# Spire.Doc Python Text Direction
## Set text direction in Word document sections and table cells
```python
# Create a new document
doc = Document()
# Add the first section
section1 = doc.AddSection()
# Set text direction for all text in a section
section1.TextDirection = TextDirection.RightToLeft
# Set Font Style and Size
style = ParagraphStyle(doc)
style.Name = "FontStyle"
style.CharacterFormat.FontName = "Arial"
style.CharacterFormat.FontSize = 15
doc.Styles.Add(style)
# Add two paragraphs and apply the font style
p = section1.AddParagraph()
p.AppendText("Only Spire.Doc, no Microsoft Office automation")
p.ApplyStyle(style.Name)
p = section1.AddParagraph()
p.AppendText("Convert file documents with high quality")
p.ApplyStyle(style.Name)
# Set text direction for a part of text
# Add the second section
section2 = doc.AddSection()
# Add a table
table = section2.AddTable()
table.ResetCells(1, 1)
cell = table.Rows[0].Cells[0]
table.Rows[0].Height = 150
table.Rows[0].Cells[0].SetCellWidth(10, CellWidthType.Point)
# Set vertical text direction of table
cell.CellFormat.TextDirection = TextDirection.RightToLeftRotated
cell.AddParagraph().AppendText("This is vertical style")
# Add a paragraph and set horizontal text direction
p = section2.AddParagraph()
p.AppendText("This is horizontal style")
p.ApplyStyle(style.Name)
```

---

# Spire.Doc Python Text Splitting
## Add columns to document section and enable line between columns
```python
#Create a new document
doc = Document()
#Add a column to the first section and set width and spacing
doc.Sections[0].AddColumn(100, 20)
#Add a line between the two columns
doc.Sections[0].PageSetup.ColumnsLineBetween = True
```

---

# Spire.Doc for Python - Alter Language Dictionary
## Set language dictionary for Word document
```python
#Create Word document.
document = Document()
#Add new section and paragraph to the document.
sec = document.AddSection()
para = sec.AddParagraph()
#Add a textRange for the paragraph and append some Peru Spanish words.
txtRange = para.AppendText("corrige según diccionario en inglés")
txtRange.CharacterFormat.LocaleIdASCII = 10250
```

---

# Spire.Doc Python Check File Format
## Detect and identify the format of a Word document
```python
# Create Word document
doc = Document()

# Load a file
doc.LoadFromFile("path_to_your_file.docx")

# Get file format
fileFormat = doc.DetectedFormatType

# Identify the format type
formatType = ""
if fileFormat == FileFormat.Doc:
    formatType = "Doc"
elif fileFormat == FileFormat.Dot:
    formatType = "Dot"
elif fileFormat == FileFormat.Docx:
    formatType = "Docx"
elif fileFormat == FileFormat.Docm:
    formatType = "Docm"
elif fileFormat == FileFormat.Dotx:
    formatType = "Dotx"
elif fileFormat == FileFormat.Dotm:
    formatType = "Dotm"
elif fileFormat == FileFormat.Rtf:
    formatType = "Rtf"
elif fileFormat == FileFormat.WordML:
    formatType = "WordML"
elif fileFormat == FileFormat.Html:
    formatType = "Html"
elif fileFormat == FileFormat.WordXml:
    formatType = "WordXml"
elif fileFormat == FileFormat.Odt:
    formatType = "Odt"
elif fileFormat == FileFormat.Ott:
    formatType = "Ott"
elif fileFormat == FileFormat.DocPre97:
    formatType = "DocPre97"
else:
    formatType = "Unknown"

# Close the document
doc.Close()
```

---

# Spire.Doc Document Comparison
## Compare two Word documents and highlight differences
```python
#Load the first document
doc1 = Document()
doc1.LoadFromFile(inputFile1)
#Load the second document
doc2 = Document()
doc2.LoadFromFile(inputFile2)
#Compare the two documents
doc1.Compare(doc2, "E-iceblue")
#Save as docx file.
doc1.SaveToFile(outputFile, FileFormat.Docx2013)
doc1.Close()
doc2.Close()
```

---

# spire.doc document comparison
## compare two Word documents with options
```python
#Set options
compareOptions = CompareOptions()
compareOptions.IgnoreFormatting = True
#Compare the two documents
doc1.Compare(doc2, "E-iceblue", DateTime.get_Now(), compareOptions)
```

---

# spire.doc python word count
## Count characters and words in a Word document
```python
# Create Word document
document = Document()
# Count the number of words
content = ""
content += "CharCount: " 
content += str(document.BuiltinDocumentProperties.CharCount)
content += "\n"
content += "CharCountWithSpace: " 
content += str(document.BuiltinDocumentProperties.CharCountWithSpace)
content += "\n"
content += "WordCount: " 
content += str(document.BuiltinDocumentProperties.WordCount)
content += "\n"
```

---

# spire.doc document properties
## set built-in document properties
```python
# Set built-in document properties
document.BuiltinDocumentProperties.Title = "Document Demo Document"
document.BuiltinDocumentProperties.Subject = "demo"
document.BuiltinDocumentProperties.Author = "James"
document.BuiltinDocumentProperties.Company = "e-iceblue"
document.BuiltinDocumentProperties.Manager = "Jakson"
document.BuiltinDocumentProperties.Category = "Doc Demos"
document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo"
document.BuiltinDocumentProperties.Comments = "This document is just a demo."
```

---

# Download Word Document from URL
## This code demonstrates how to download a Word document from a URL and save it locally
```python
outputFile = "DownloadWordFromURL.docx"

# Create Word document
document = Document()

# Download Word document from URL
response = requests.get("https://www.e-iceblue.com/images/test.docx")
if response.status_code == 200:
    content = response.content
    ms = Stream(content)
    document.LoadFromStream(ms, FileFormat.Docx)

# Save to file
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
```

---

# Spire.Doc Python Document Properties
## Extract builtin and custom document properties from a Word document
```python
inputFile = "data/Properties.docx"

# Create a document
document = Document()

# Load the document from disk
document.LoadFromFile(inputFile)

# Get Builtin document properties
title = document.BuiltinDocumentProperties.Title
comments = document.BuiltinDocumentProperties.Comments
author = document.BuiltinDocumentProperties.Author
keywords = document.BuiltinDocumentProperties.Keywords
company = document.BuiltinDocumentProperties.Company

# Get custom document properties
for i in range(document.CustomDocumentProperties.Count):
    propertyName = document.CustomDocumentProperties[i].Name
    propertyValue = document.CustomDocumentProperties.get_Item(i).ToString()
```

---

# spire.doc document operations
## Load document from disk and save to disk
```python
inputFile = "./Data/Sample.docx"
outputFile = "LoadAndSaveToDisk.docx"

#Create a new document
doc = Document()
# Load the document from the absolute/relative path on disk.
doc.LoadFromFile(inputFile)
# Save the document to disk
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()
```

---

# spire.doc document stream operations
## load document from stream and save to stream
```python
# Open the stream. Read only access is enough to load a document.
stream = Stream("input_file_path")
# Load the entire document into memory.
doc = Document(stream)
# You can close the stream now, it is no longer needed because the document is in memory.
stream.Close()
# Do something with the document
# Convert the document to a different format and save to stream.
newStream = Stream()
doc.SaveToStream(newStream, FileFormat.Rtf)
# Rewind the stream position back to zero so it is ready for the next reader.
newStream.Position = 0
doc.Close()
```

---

# Spire.Doc Python Document Object Traversal
## Recursively traverse and collect information about all document objects in a Word document
```python
builder = ""
#find all document object
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    SectionIndex = document.GetIndex(section)
    builder += "section index {} has following ChildObjects".format(SectionIndex)
    builder += "\n"
    for j in range(section.Body.ChildObjects.Count):
        obj = section.Body.ChildObjects.get_Item(j)
        builder += "Index : {}, ChildObject Type: {}".format(section.Body.GetIndex(obj), obj.DocumentObjectType.name)
        builder += "\n"
        
        if obj.DocumentObjectType is DocumentObjectType.Paragraph:
            paragraph = obj if isinstance(obj, Paragraph) else None
            builder += "\tParagraph index {} has following ChildObjects".format(section.Body.GetIndex(paragraph))
            builder += "\n"
            for k in range(paragraph.ChildObjects.Count):
                obj2 = paragraph.ChildObjects.get_Item(k)
                builder += "\tIndex : {}, ChildObject Type: {}".format(paragraph.GetIndex(obj2), obj2.DocumentObjectType.name)
                builder += "\n"
    builder += " "
    builder += "\n"
```

---

# Spire.Doc Document Properties
## Set built-in and custom document properties in Word documents
```python
#Set the build-in Properties.
document.BuiltinDocumentProperties.Title = "Document Demo Document"
document.BuiltinDocumentProperties.Author = "James"
document.BuiltinDocumentProperties.Company = "e-iceblue"
document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo"
document.BuiltinDocumentProperties.Comments = "This document is just a demo."

#Set the custom properties.
custom = document.CustomDocumentProperties
custom.Add("e-iceblue", Boolean(True))
custom.Add("Authorized By", String("John Smith"))
custom.Add("Authorized Date", DateTime.get_Today())
```

---

# Spire.Doc Python Document View Modes
## Set view modes for Word documents
```python
# Set Word view modes.
document.ViewSetup.DocumentViewType = DocumentViewType.WebLayout
document.ViewSetup.ZoomPercent = 150
document.ViewSetup.ZoomType = ZoomType.none
```

---

# spire.doc document operation
## update last saved date
```python
def LocalTimeToGreenwishTime(localTime):
    localTimeZone = TimeZone.get_CurrentTimeZone()
    timeSpan = localTimeZone.GetUtcOffset(localTime)
    greenwishTime = localTime - timeSpan
    return greenwishTime

# Create Word document
document = Document()
# Update the last saved date
document.BuiltinDocumentProperties.LastSaveDate = LocalTimeToGreenwishTime(DateTime.get_Now())
```

---

# Spire.Doc Document Section Append
## Add a section from one Word document to another
```python
#Open a Word document as target document
TarDoc = Document(inputFile1)
#Open a Word document as source document
SouDoc = Document(inputFile2)
#Get the second section from source document
Ssection = SouDoc.Sections[0]
#Add the section in target document
TarDoc.Sections.Add(Ssection.Clone())
#Save the file
TarDoc.SaveToFile(outputFile, FileFormat.Docx)
SouDoc.Close()
TarDoc.Close()
```

---

# spire.doc python document cloning
## clone a Word document using spire.doc library
```python
#Create Word document.
document = Document()
#Clone the word document.
newDoc = document.Clone()
```

---

# Spire.Doc Python Document Content Copying
## Copy content from one document to another document
```python
#Copy content from source file and insert them to the target file.
for i in range(sourceDoc.Sections.Count):
    sec = sourceDoc.Sections.get_Item(i)
    for j in range(sec.Body.ChildObjects.Count):
        obj = sec.Body.ChildObjects.get_Item(j)
        destinationDoc.Sections[0].Body.ChildObjects.Add(obj.Clone())     
```

---

# Spire.Doc Python Document Operation
## Keep same format when appending documents
```python
#Create document objects
srcDoc = Document()
destDoc = Document()
#Keep same format of source document
srcDoc.KeepSameFormat = True
#Copy the sections of source document to destination document
for i in range(srcDoc.Sections.Count):
    section = srcDoc.Sections.get_Item(i)
    destDoc.Sections.Add(section.Clone())
```

---

# Spire.Doc Python Headers and Footers Linking
## Link headers and footers between documents and clone sections
```python
# Link the headers and footers in the source file
srcDoc.Sections[0].HeadersFooters.Header.LinkToPrevious = True
srcDoc.Sections[0].HeadersFooters.Footer.LinkToPrevious = True
# Clone the sections of source to destination
for i in range(srcDoc.Sections.Count):
    section = srcDoc.Sections.get_Item(i)
    dstDoc.Sections.Add(section.Clone())
```

---

# Spire.Doc Python Document Merge
## Demonstrates how to merge two Word documents by adding sections from one document to another
```python
#Create word document
document = Document()
document.LoadFromFile(inputFile1, FileFormat.Docx)
#Load second file 
documentMerge = Document()
documentMerge.LoadFromFile(inputFile2, FileFormat.Docx)
#merge
for i in range(documentMerge.Sections.Count):
    sec = documentMerge.Sections.get_Item(i)
    document.Sections.Add(sec.Clone())
```

---

# Spire.Doc Python Document Merge
## Merge documents on the same page
```python
#Create a document
document = Document()
#Clone a destination document
destinationDocument = Document()
#Traverse sections
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    #Traverse body ChildObjects
    for j in range(section.Body.ChildObjects.Count):
        obj = section.Body.ChildObjects.get_Item(j)
        #Clone to destination document at the same page
        destinationDocument.Sections[0].Body.ChildObjects.Add(obj.Clone())
#Save the document.
destinationDocument.SaveToFile("MergedDocument.docx", FileFormat.Docx)
document.Close()
destinationDocument.Close()
```

---

# Spire.Doc Python Document Operation
## Preserve theme when appending documents
```python
# Create source and destination documents
doc = Document()
newWord = Document()
# Clone style, theme, and compatibility from source to destination
doc.CloneDefaultStyleTo(newWord)
doc.CloneThemesTo(newWord)
doc.CloneCompatibilityTo(newWord)
# Add section from source to destination
newWord.Sections.Add(doc.Sections[0].Clone())
```

---

# Spire.Doc Python Section Break
## Set section breaks to continuous in a Word document
```python
# Iterate through all sections in the document
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    # Set section break as continuous
    section.BreakCode = SectionBreakType.NoBreak
```

---

# Spire.Doc document insertion
## Insert one Word document into another
```python
# Load the Word document
doc = Document()
doc.LoadFromFile(inputFile1)
# Insert document from file
doc.InsertTextFromFile(inputFile2, FileFormat.Auto)
```

---

# Spire.Doc Document Splitting by Page Breaks
## Split a Word document into multiple documents at each page break
```python
#Create Word document.
original = Document()
#Create a new word document and add a section to it.
newWord = Document()
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
```

---

# Spire.Doc Document Splitting
## Split Word document into multiple HTML pages based on Heading1 styles
```python
def IsInNextDocument(element):
    if isinstance(element, Paragraph):
        p = element if isinstance(element, Paragraph) else None
        if p.StyleName == "Heading1":
            return True
    return False

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
                #Embed css style and image data into html page
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
    #Embed css style and image data into html page
    subDoc.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal
    subDoc.HtmlExportOptions.ImageEmbedded = True
    #Save to html file
    resultF = outputFolder+"out-{0}.html".format(index)
    subDoc.SaveToFile(resultF, FileFormat.Html)
    index += 1
    subDoc.Close()
```

---

# Spire.Doc Python Track Changes
## Accept or reject tracked changes in a Word document
```python
#Get the first section and the paragraph we want to accept/reject the changes.
sec = document.Sections[0]
para = sec.Paragraphs[0]
#Accept the changes or reject the changes.
para.Document.AcceptChanges()
#para.Document.RejectChanges()
```

---

# Spire.Doc Document Revision Comparison
## Compare two documents and extract revision differences
```python
# Compare two documents
doc1.Compare(doc2, "Author")
revisions = DifferRevisions(doc1)

# Get insert revision
insertRevisionList = revisions.InsertRevisions

# Get deleted revisions
deleteRevisionList = revisions.DeleteRevisions

# Process insert revisions
for i in range(0, insertRevisionList.__len__()):
    if insertRevisionList[i].DocumentObjectType == DocumentObjectType.TextRange:
        textRange = TextRange(insertRevisionList[i])
        content += "insert #" + str(i+1) + ":" + textRange.Text + '\n'
        content += "=====================" + '\n'

# Process delete revisions
for i in range(0, deleteRevisionList.__len__()):
    if deleteRevisionList[i].DocumentObjectType == DocumentObjectType.TextRange:
        textRange = TextRange(deleteRevisionList[i])
        content += "delete #" + str(i+1) + ":" + textRange.Text + '\n'
        content += "=====================" + '\n'
```

---

# spire.doc python track changes
## Enable track changes in a Word document
```python
# Create Word document
document = Document()
# Enable the track changes
document.TrackChanges = True
```

---

# Spire.Doc Python Document Revisions
## Extract insert and delete revisions from a Word document
```python
#Create Word document.
document = Document()
document.LoadFromFile(inputFile)
#Traverse sections
for k in range(document.Sections.Count):
    sec = document.Sections.get_Item(k)
    #Iterate through the element under body in the section
    for m in range(sec.Body.ChildObjects.Count):
        docItem = sec.Body.ChildObjects.get_Item(m)
        if isinstance(docItem, Paragraph):
            para = docItem
            #Determine if the paragraph is an insertion revision
            if para.IsInsertRevision:
                #Get insertion revision
                insRevison = para.InsertRevision
                #Get insertion revision type
                insType = insRevison.Type
                #Get insertion revision author
                insAuthor = insRevison.Author
            #Determine if the paragraph is a delete revision
            elif para.IsDeleteRevision:
                delRevison = para.DeleteRevision
                delType = delRevison.Type
                delAuthor = delRevison.Author
            #Iterate through the element in the paragraph
            for j in range(para.ChildObjects.Count):
                obj = para.ChildObjects.get_Item(j)
                if isinstance(obj, TextRange):
                    textRange = obj
                    #Determine if the textrange is an insertion revision
                    if textRange.IsInsertRevision:
                        insRevison = textRange.InsertRevision
                        insType = insRevison.Type
                        insAuthor = insRevison.Author
                    elif textRange.IsDeleteRevision:
                        #Determine if the textrange is a delete revision
                        delRevison = textRange.DeleteRevision
                        delType = delRevison.Type
                        delAuthor = delRevison.Author
```

---

# spire.doc python revision time modification
## modify revision timestamps in document
```python
# Specify the date string and format
dateString = "2023/3/1 00:00:00"
formatStr = "yyyy/M/d HH:mm:ss"

# Parse the date string into a DateTime object using the specified format
date = DateTime.ParseExact(dateString, formatStr)

# Iterate through the sections in the document
for i in range(document.Sections.Count):
    sec = document.Sections[i]
    # Iterate through the child objects in the section's body
    for j in range(sec.Body.ChildObjects.Count):
        docItem = sec.Body.ChildObjects.get_Item(j)
        # Check if the child object is a Paragraph
        if isinstance(docItem, Paragraph):
            # Cast the child object to a Paragraph
            para = docItem

            # Check if the paragraph contains an insert revision
            if para.IsInsertRevision:
                # Get the InsertRevision object for the paragraph
                insRevison = para.InsertRevision

                # Set the DateTime property of the insert revision to the specified date
                insRevison.DateTime = date
            # Check if the paragraph contains a delete revision
            elif para.IsDeleteRevision:
                # Get the DeleteRevision object for the paragraph
                delRevison = para.DeleteRevision

                # Set the DateTime property of the delete revision to the specified date
                delRevison.DateTime = date

            # Iterate through the child objects in the paragraph
            for k in range(para.ChildObjects.Count):
                obj = para.ChildObjects.get_Item(k)
                # Check if the child object is a TextRange
                if isinstance(obj, TextRange):
                    # Cast the child object to a TextRange
                    textRange = obj

                    # Check if the text range contains an insert revision
                    if textRange.IsInsertRevision:
                        # Get the InsertRevision object for the text range
                        insRevison = textRange.InsertRevision

                        # Set the DateTime property of the insert revision to the specified date
                        insRevison.DateTime = date
                    # Check if the text range contains a delete revision
                    elif textRange.IsDeleteRevision:
                        # Get the DeleteRevision object for the text range
                        delRevison = textRange.DeleteRevision

                        # Set the DateTime property of the delete revision to the specified date
                        delRevison.DateTime = date
```

---

# spire.doc python variables
## add document variables to Word document
```python
#Create Word document.
document = Document()
#Add a section.
section = document.AddSection()
#Add a paragraph.
paragraph = section.AddParagraph()
#Add a DocVariable field.
paragraph.AppendField("A1", FieldType.FieldDocVariable)
#Add a document variable to the DocVariable field.
document.Variables.Add("A1", "12")
#Update fields.
document.IsUpdateFields = True
```

---

# Spire.Doc Python Variables
## Count variables in a Word document
```python
# Create Word document
document = Document()
# Load the file from disk
document.LoadFromFile(inputFile)
# Get the number of variables in the document
number = document.Variables.Count
```

---

# Spire.Doc Python Variables
## Extract variables from a Word document
```python
# Extract variables from a loaded document
stringBuilder = ""
stringBuilder += "This document has following variables:"
stringBuilder += "\n"
for i in range(document.Variables.Count):
    name = document.Variables.GetNameByIndex(i)
    value = document.Variables.GetValueByIndex(i)
    stringBuilder += "Name: " 
    stringBuilder += name 
    stringBuilder += ", "
    stringBuilder += "Value: " 
    stringBuilder += value
    stringBuilder += "\n"
```

---

# Spire.Doc Python Variables
## Remove variables from Word document
```python
#Remove the variable by name.
document.Variables.Remove("A1")
document.IsUpdateFields = True
```

---

# spire.doc python variables
## retrieve document variables
```python
#Create Word document
document = Document()
#Load the file from disk
document.LoadFromFile(inputFile)
#Retrieve name of the variable by index
s1 = document.Variables.GetNameByIndex(0)
#Retrieve value of the variable by index
s2 = document.Variables.GetValueByIndex(0)
#Retrieve the value of the variable by name
s3 = document.Variables["A1"]
document.Close()
```

---

# Spire.Doc Python Gradient Background
## Set gradient background for Word document
```python
#Create Word document.
document = Document()
#Set the background type as Gradient.
document.Background.Type = BackgroundType.Gradient
Test = document.Background.Gradient
#Set the first color and second color for Gradient.
Test.Color1 = Color.get_White()
Test.Color2 = Color.get_LightBlue()
#Set the Shading style and Variant for the gradient.
Test.ShadingVariant = GradientShadingVariant.ShadingDown
Test.ShadingStyle = GradientShadingStyle.Horizontal
```

---

# Spire.Doc Python Document Background
## Set image background for Word document
```python
# Set the background type as picture
document.Background.Type = BackgroundType.Picture
# Set the background picture
document.Background.SetPicture(inputFile_Img)
```

---

# Spire.Doc Page Setup
## Add gutter to Word document section
```python
#Create Word document
document = Document()
#Create a new section
section = document.Sections[0]
#Set gutter
section.PageSetup.Gutter = 100
```

---

# Spire.Doc Python Line Numbering
## Add line numbers to Word document
```python
#Set the start value of the line numbers.
document.Sections[0].PageSetup.LineNumberingStartValue = 1
#Set the interval between displayed numbers.
document.Sections[0].PageSetup.LineNumberingStep = 6
#Set the distance between line numbers and text.
document.Sections[0].PageSetup.LineNumberingDistanceFromText = 40
#Set the numbering mode of line numbers. There are four choices: None, Continuous, RestartPage and RestartSection.
document.Sections[0].PageSetup.LineNumberingRestartMode = LineNumberingRestartMode.Continuous
```

---

# Spire.Doc Python Page Borders
## Add page borders to a Word document with custom style, color, and spacing
```python
#Add page borders with special style and color.
document.Sections[0].PageSetup.Borders.BorderType(BorderStyle.DoubleWave)
document.Sections[0].PageSetup.Borders.Color(Color.get_LightSeaGreen())

#Set the space between border and text.
document.Sections[0].PageSetup.Borders.Left.Space = 50
document.Sections[0].PageSetup.Borders.Right.Space = 50
```

---

# spire.doc python page setup
## add page numbers in document sections
```python
# Repeat step2 and Step3 for the rest sections, so change the code with for loop.
for i in range(0, 3):
    footer = document.Sections[i].HeadersFooters.Footer
    footerParagraph = footer.AddParagraph()
    footerParagraph.AppendField("page number", FieldType.FieldPage)
    footerParagraph.AppendText(" of ")
    footerParagraph.AppendField(
        "number of pages", FieldType.FieldSectionPages)
    footerParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

    if i == 2:
         break
    else:
        document.Sections[i + 1].PageSetup.RestartPageNumbering = True
        document.Sections[i + 1].PageSetup.PageStartingNumber = 1
```

---

# Spire.Doc Python Page Setup
## Configure different page settings for document sections
```python
# Get the second section 
SectionTwo = doc.Sections[1]

# Set the orientation
SectionTwo.PageSetup.Orientation = PageOrientation.Landscape

# Set page size
#SectionTwo.PageSetup.PageSize = new SizeF(800, 800)
```

---

# spire.doc python break insertion
## insert section break in word document
```python
#insert a break code
section = document.AddSection()
section.AddParagraph().InsertSectionBreak(SectionBreakType.NewPage)
```

---

# Spire.Doc Python Page Break
## Insert page breaks after specified text in a Word document
```python
#Find the specified word "technology" where we want to insert the page break.
selections = document.FindAllString("technology", True, True)
#Traverse each word "technology".
for ts in selections:
    range = ts.GetAsOneRange()
    paragraph = range.OwnerParagraph
    index = paragraph.ChildObjects.IndexOf(range)
    #Create a new instance of page break and insert a page break after the word "technology".
    pageBreak = Break(document, BreakType.PageBreak)
    paragraph.ChildObjects.Insert(index + 1, pageBreak)
```

---

# Spire.Doc Python Page Break
## Insert page break at specific paragraph
```python
# Insert page break
document.Sections[0].Paragraphs[3].AppendBreak(BreakType.PageBreak)
```

---

# spire.doc python section break
## insert section break in Word document
```python
#Create Word document.
document = Document()
#Insert section break. There are five section break options including EvenPage, NewColumn, NewPage, NoBreak, OddPage.
document.Sections[0].Paragraphs[1].InsertSectionBreak(SectionBreakType.NoBreak)
```

---

# Spire.Doc Page Setup
## Configure page settings, headers and footers in Word document
```python
#Create Word document.
document = Document()
section = document.AddSection()
#The unit of all measures below is point, 1point = 0.3528 mm.
section.PageSetup.PageSize = PageSize.A4()
section.PageSetup.Margins.Top = 72
section.PageSetup.Margins.Bottom = 72
section.PageSetup.Margins.Left = 89.85
section.PageSetup.Margins.Right = 89.85
#Insert header and footer.
header = section.HeadersFooters.Header
footer = section.HeadersFooters.Footer
#Insert picture and text to header.
headerParagraph = header.AddParagraph()
headerPicture = headerParagraph.AppendPicture("./Data/Header.png")
#Header text.
text = headerParagraph.AppendText("Demo of Spire.Doc")
text.CharacterFormat.FontName = "Arial"
text.CharacterFormat.FontSize = 10
text.CharacterFormat.Italic = True
headerParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right
#Border.
headerParagraph.Format.Borders.Bottom.BorderType = BorderStyle.Single
headerParagraph.Format.Borders.Bottom.Space = 0.05
#Header picture layout - text wrapping.
headerPicture.TextWrappingStyle = TextWrappingStyle.Behind
#Header picture layout - position.
headerPicture.HorizontalOrigin = HorizontalOrigin.Page
headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
headerPicture.VerticalOrigin = VerticalOrigin.Page
headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top
#Insert picture to footer.
footerParagraph = footer.AddParagraph()
footerPicture = footerParagraph.AppendPicture("./Data/Footer.png")
#Footer picture layout.
footerPicture.TextWrappingStyle = TextWrappingStyle.Behind
footerPicture.HorizontalOrigin = HorizontalOrigin.Page
footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
footerPicture.VerticalOrigin = VerticalOrigin.Page
footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom
#Insert page number.
footerParagraph.AppendField("page number", FieldType.FieldPage)
footerParagraph.AppendText(" of ")
footerParagraph.AppendField("number of pages", FieldType.FieldNumPages)
footerParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right
#Border.
footerParagraph.Format.Borders.Top.BorderType = BorderStyle.Single
footerParagraph.Format.Borders.Top.Space = 0.05
```

---

# Spire.Doc Python Remove Page Breaks
## Core functionality to remove page breaks from a Word document
```python
#Traverse every paragraph of the first section of the document.
for j in range(document.Sections[0].Paragraphs.Count):
    p = document.Sections[0].Paragraphs[j]
    #Traverse every child object of a paragraph.
    for i in range(p.ChildObjects.Count):
        obj = p.ChildObjects[i]
        #Find the page break object.
        if obj.DocumentObjectType == DocumentObjectType.Break:
            b = obj if isinstance(obj, Break) else None
            #Remove the page break object from paragraph.
            p.ChildObjects.Remove(b)
```

---

# Spire.Doc Page Number Reset
## Reset page numbering in Word documents by combining multiple documents and modifying page number fields
```python
# Use section method to combine all documents into one word document
for i in range(document2.Sections.Count):
    sec = document2.Sections.get_Item(i)
    document1.Sections.Add(sec.Clone())
for i in range(document3.Sections.Count):
    sec = document3.Sections.get_Item(i)
    document1.Sections.Add(sec.Clone())

# Traverse every section of document1
for i in range(document1.Sections.Count):
    sec = document1.Sections.get_Item(i)
    # Traverse every object of the footer
    for j in range(sec.HeadersFooters.Footer.ChildObjects.Count):
        obj = sec.HeadersFooters.Footer.ChildObjects.get_Item(j)
        if obj.DocumentObjectType == DocumentObjectType.StructureDocumentTag:
            para = obj.ChildObjects[0]
            for k in range(para.ChildObjects.Count):
                item = para.ChildObjects.get_Item(k)
                if item.DocumentObjectType == DocumentObjectType.Field:
                    # Find the item and its field type is FieldNumPages
                    if ( item if isinstance(item, Field) else None).Type == FieldType.FieldNumPages:
                        # Change field type to FieldSectionPages
                        ( item if isinstance(item, Field) else None).Type = FieldType.FieldSectionPages

# Restart page number of section and set the starting page number to 1
document1.Sections[1].PageSetup.RestartPageNumbering = True
document1.Sections[1].PageSetup.PageStartingNumber = 1
document1.Sections[2].PageSetup.RestartPageNumbering = True
document1.Sections[2].PageSetup.PageStartingNumber = 1
```

---

# Spire.Doc Python Gutter Position
## Set gutter position in Word document
```python
# Get the first section of the document.
section = document.Sections[0]

# Set the top gutter option to true for the section's page setup.
section.PageSetup.IsTopGutter = True

# Set the width of the gutter in points (100f).
section.PageSetup.Gutter = 100
```

---

# Spire.Doc Document to Bytes Conversion
## Convert a Word document to bytes and back to a document object
```python
# Create a new memory stream.
outStream = Stream()
# Save the document to stream.
doc.SaveToStream(outStream, FileFormat.Docx)
# Convert the document to bytes.
docBytes = outStream.ToArray()
# The bytes are now ready to be stored/transmitted.
# Now reverse the steps to load the bytes back into a document object.
inStream = Stream(docBytes)
# Load the stream into a new document object.
newDoc = Document(inStream)
```

---

# spire.doc python object to image conversion
## Convert document objects to images

```python
def _ConvertParagraphToImage(obj):
    doc = Document()
    section = doc.AddSection()

    section.Body.ChildObjects.Add(obj.Clone())
    stream = doc.SaveImageToStreams(0, ImageType.Bitmap)
    doc.Close()
    return stream


def _ConvertTableToImage(obj):
    doc = Document()
    section = doc.AddSection()

    section.Body.ChildObjects.Add(obj.Clone())

    stream = doc.SaveImageToStreams(0, ImageType.Bitmap)
    doc.Close()
    return stream


def _ConvertTableRowToImage(obj):
    doc = Document()
    section = doc.AddSection()
    table = section.AddTable()
    table.Rows.Add(obj.Clone())
    stream = doc.SaveImageToStreams(0, ImageType.Bitmap)
    doc.Close()
    return stream


def _ConvertTableCellToImage(obj):
    doc = Document()
    section = doc.AddSection()
    table = section.AddTable()
    table.AddRow().Cells.Add(obj.Clone())
    stream = doc.SaveImageToStreams(0, ImageType.Bitmap)
    doc.Close()
    return stream


def _ConvertShapeToImage(obj):
    doc = Document()
    section = doc.AddSection()
    section.AddParagraph().ChildObjects.Add(obj.Clone())
    ms = Stream()
    doc.SaveToStream(ms, FileFormat.Docx)
    doc.LoadFromStream(ms, FileFormat.Docx)
    stream = doc.SaveImageToStreams(0, ImageType.Bitmap)
    ms.Close()
    doc.Close()
    return stream
```

---

# spire.doc python conversion
## convert markdown to various document formats
```python
# Create a document object
document = Document()
# Load markdown file
document.LoadFromFile("Data/FromMarkdown.md")
# Save to markdown format
document.SaveToFile("FromMarkdown_markdown.md", FileFormat.Markdown)
# Save to docx format
document.SaveToFile("FromMarkdown_docx.docx", FileFormat.Docx)
# Save to doc format
document.SaveToFile("FromMarkdown_doc.doc", FileFormat.Doc)
# Save to pdf format
document.SaveToFile("FromMarkdown_pdf.pdf", FileFormat.PDF)
```

---

# HTML to Image Conversion
## Convert HTML document to image using Spire.Doc
```python
# Create Word document
document = Document()
# Load the HTML file
document.LoadFromFile(inputFile, FileFormat.Html, XHTMLValidationType.none)
# Convert document to image stream
imageStream = document.SaveImageToStreams(0, ImageType.Bitmap)
# Save image stream to file
with open(outputFile,'wb') as imageFile:
    imageFile.write(imageStream.ToArray())
document.Close()
```

---

# Spire.Doc HTML to PDF Conversion
## Convert HTML files to PDF format using Spire.Doc library

```python
# Basic HTML to PDF conversion
document = Document()
document.LoadFromFile(inputFile, FileFormat.Html, XHTMLValidationType.none)
document.SaveToFile(outputFile, FileFormat.PDF)
document.Close()

# HTML to PDF conversion with PostScript parameters
document = Document()
document.LoadFromFile(inputFile, FileFormat.Html, XHTMLValidationType.none)
Ps = ToPdfParameterList()
Ps.UsePSCoversion = True
document.SaveToFile(outputFile, Ps)
document.Close()
```

---

# spire.doc python html to xml conversion
## convert html file to xml format
```python
#Create Word document
document = Document()
#Load the file from disk
document.LoadFromFile(inputFile)
#Save to file
document.SaveToFile(outputFile, FileFormat.Xml)
document.Close()
```

---

# spire.doc python conversion
## convert HTML to XPS format
```python
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile, FileFormat.Html, XHTMLValidationType.none)
#Save to file.
document.SaveToFile(outputFile, FileFormat.XPS)
document.Close()
```

---

# Spire.Doc Image to PDF Conversion
## Convert an image to PDF document using Spire.Doc library
```python
# Create a new document
doc = Document()
# Create a new section
section = doc.AddSection()
# Create a new paragraph
paragraph = section.AddParagraph()
# Add a picture for paragraph
picture = paragraph.AppendPicture(inputFile)
# Set A4 page size
section.PageSetup.PageSize = PageSize.A4()
# Set the page margins
section.PageSetup.Margins.Top = 10
section.PageSetup.Margins.Left = 25
doc.SaveToFile(outputFile, FileFormat.PDF)
doc.Close()
```

---

# spire.doc document conversion
## convert ODT file to Word document
```python
inputFile =  "./Data/Template_OdtFile.odt"
outputFile = "OdtToWord.docx"
#Create Word document
document = Document()
#Load the file from disk
document.LoadFromFile(inputFile)
#Save to Docx file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
```

---

# RTF to HTML Conversion
## Convert RTF document to HTML format using Spire.Doc
```python
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Save to file.
document.SaveToFile(outputFile, FileFormat.Html)
document.Close()
```

---

# RTF to PDF Conversion
## Convert RTF document to PDF format using Spire.Doc
```python
# Define input and output file paths
inputFile = "./Data/Template_RtfFile.rtf"
outputFile = "RtfToPDF.pdf"

# Create a document object
doc = Document()
# Load RTF file
doc.LoadFromFile(inputFile)
# Save as PDF
doc.SaveToFile(outputFile, FileFormat.PDF)
# Close document
doc.Close()
```

---

# Spire.Doc document to image conversion
## Convert Word document to image format
```python
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

---

# Spire.Doc Python Conversion
## Convert Word document to Markdown format
```python
# Create a Document object
doc = Document()

# Load a Word document
doc.LoadFromFile("Data/ToMarkdown.docx")

# Convert to Markdown format
doc.SaveToFile("ToMarkdown_output.md", FileFormat.Markdown)
```

---

# spire.doc python conversion
## convert doc to odt format
```python
# Create word document
document = Document()
document.LoadFromFile(inputFile)
# Save doc file to ODT format
document.SaveToFile(outputFile, FileFormat.Odt)
document.Close()
```

---

# Spire.Doc Python Conversion
## Convert DOCX file to PCL format
```python
# Create a new document
doc = Document()
# Load from file
doc.LoadFromFile(inputFile)
# Save to PCL format
doc.SaveToFile(outputFile, FileFormat.PCL)
# Close the document
doc.Close()
```

---

# Spire.Doc Python Conversion
## Convert document to PostScript format
```python
# Create document object
doc = Document()
# Load input file
doc.LoadFromFile(inputFile)
# Save to PostScript format
doc.SaveToFile(outputFile, FileFormat.PostScript)
# Close document
doc.Close()
```

---

# spire.doc document conversion
## convert document to RTF format
```python
#Create word document
document = Document()
document.LoadFromFile(inputFile)
#Save doc file.
document.SaveToFile(outputFile, FileFormat.Rtf)
document.Close()
```

---

# Spire.Doc Python Document Conversion
## Convert Word document to SVG format
```python
inputFile = "./Data/ToSVGTemplate.docx"
outputFile = "ToSVG.svg"
#Create word document
document = Document()
document.LoadFromFile(inputFile)
document.SaveToFile(outputFile, FileFormat.SVG)
document.Close()
```

---

# Spire.Doc Document Conversion
## Convert Word document to XML format
```python
#Create word document.
document = Document()
document.LoadFromFile(inputFile)
#Save the document to a xml file.
document.SaveToFile(outputFile, FileFormat.Xml)
document.Close()
```

---

# Spire.Doc document conversion
## Convert Word document to XPS format
```python
# Define input and output files
inputFile = "./Data/ConvertedTemplate.docx"
outputFile = "ToXPS.xps"
# Create word document
document = Document()
document.LoadFromFile(inputFile)
# Save the document to a xps file
document.SaveToFile(outputFile, FileFormat.XPS)
document.Close()
```

---

# spire.doc text to word conversion
## convert text file to word document
```python
#Create Word document
document = Document()
#Load the file from disk
document.LoadFromFile(inputFile)
#Save the file
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
```

---

# Word to PDF/A Conversion
## Convert Word document to PDF/A format using Spire.Doc
```python
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Set the Conformance-level of the Pdf file to PDF_A1B.
toPdf = ToPdfParameterList()
toPdf.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B
#Save the file.
document.SaveToFile(outputFile, toPdf)
document.Close()
```

---

# Spire.Doc Python Word to Text Conversion
## Convert Word document to text format
```python
# Create Word document
document = Document()
# Load the file from disk
document.LoadFromFile(inputFile)
# Save as text file
document.SaveToFile(outputFile, FileFormat.Txt)
document.Close()
```

---

# spire.doc python conversion
## convert Word document to Word XML formats
```python
#Create Word document
document = Document()

#Load the file from disk
document.LoadFromFile(inputFile)

#For word 2003
document.SaveToFile(outputFile_2003, FileFormat.WordML)

#For word 2007
document.SaveToFile(outputFile_2007, FileFormat.WordXml)
document.Close()
```

---

# Spire.Doc XML to PDF Conversion
## Convert XML file to PDF format
```python
#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile, FileFormat.Xml)

#Save to file.
document.SaveToFile(outputFile, FileFormat.PDF)
document.Close()
```

---

# Spire.Doc XML to Word Conversion
## Convert XML file to Word document format
```python
#Create Word document
document = Document()

#Load the file from disk
document.LoadFromFile(inputFile, FileFormat.Xml)

#Save to file
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
```

---

# HTML to Word Conversion
## Convert HTML file to Word document using Spire.Doc
```python
inputFile = "./Data/InputHtmlFile.html"
outputFile = "HtmlFileToWord.docx"

#Open an html file.
document = Document()
document.LoadFromFile(inputFile, FileFormat.Html, XHTMLValidationType.none)
#Save it to a Word document.
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
```

---

# Spire.Doc Python HTML to Word Conversion
## Convert HTML string to Word document
```python
# Get html string
with open(inputFile) as fp:
    HTML = fp.read()
# Create a new document
document = Document()
# Add a section
sec = document.AddSection()
# Add a paragraph and append html string
sec.AddParagraph().AppendHTML(HTML)
# Save it to a Word file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
```

---

# spire.doc python epub conversion
## add cover image to epub conversion
```python
# Load document
doc = Document()
doc.LoadFromFile(inputFile)
# Create picture object and load cover image
picture = DocPicture(doc)
picture.LoadImage("./Data/Cover.png")
# Save to EPUB with cover image
doc.SaveToEpub(outputFile, picture)
doc.Close()
```

---

# spire.doc document conversion
## convert Word document to EPUB format
```python
inputFile = "./Data/ToEpub.doc"
outputFile = "ToEpub.epub"
doc = Document()
doc.LoadFromFile(inputFile)
doc.SaveToFile(outputFile, FileFormat.EPub)
doc.Close()
```

---

# Spire.Doc document conversion
## Convert Word document to HTML format
```python
# Create word document
document = Document()
document.LoadFromFile(inputFile)
# Save doc file.
document.SaveToFile(outputFile, FileFormat.Html)
document.Close()
```

---

# Spire.Doc HTML Export Options
## Configure HTML export options when converting Word documents to HTML
```python
#Set whether the css styles are embeded or not. 
document.HtmlExportOptions.CssStyleSheetFileName = "sample.css"
document.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.External
#Set whether the images are embeded or not. 
document.HtmlExportOptions.ImageEmbedded = False
document.HtmlExportOptions.ImagesPath = "./"
#Set the option whether to export form fields as plain text or not.
document.HtmlExportOptions.IsTextInputFormFieldAsText = True
```

---

# spire.doc python conversion
## disable hyperlinks when converting docx to pdf
```python
#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Create an instance of ToPdfParameterList.
pdf = ToPdfParameterList()
#Set DisableLink to true to remove the hyperlink effect for the result PDF page. 
#Set DisableLink to false to preserve the hyperlink effect for the result PDF page.
pdf.DisableLink = True
#Save to file.
document.SaveToFile(outputFile, pdf)
document.Close()
```

---

# Spire.Doc PDF Font Embedding
## Embed all fonts when converting Word documents to PDF
```python
#embeds full fonts by default when IsEmbeddedAllFonts is set to true.
ppl = ToPdfParameterList()
ppl.IsEmbeddedAllFonts = True
```

---

# Spire.Doc Font Embedding
## Embed non-installed fonts when converting DOC to PDF
```python
# Embed the non-installed fonts
parms = ToPdfParameterList()
fonts = []
fonts.append(PrivateFontPath("Font Name", "Path_To_Font.ttf"))
parms.PrivateFontPaths = fonts
```

---

# Spire.Doc PDF Conversion with Hidden Text
## Convert Word document to PDF while preserving hidden text
```python
# When converting to PDF file, set the property IsHidden as true
pdf = ToPdfParameterList()
pdf.IsHidden = True
```

---

# Spire.Doc Python Image Quality Setting
## Set image quality when converting Word to PDF
```python
# Create Word document
document = Document()
# Set the output image quality to be 40% of the original image. The default set of the output image quality is 80% of the original.
document.JPEGQuality = 40
```

---

# Spire.Doc PDF Conversion with Embedded Font
## Specify embedded font when converting Word document to PDF
```python
# Specify embedded font
parms = ToPdfParameterList()
part = []
part.append("PT Serif Caption")
parms.EmbeddedFontNameList = part
```

---

# Spire.Doc Python Conversion
## Convert Word document to PDF
```python
#Create word document
document = Document()
document.LoadFromFile(inputFile)
#Save the document to a PDF file.
document.SaveToFile(outputFile, FileFormat.PDF)
document.Close()
```

---

# spire.doc python conversion
## convert Word to PDF and create bookmarks
```python
document = Document()
#Load the document from disk
document.LoadFromFile(inputFile)
parames = ToPdfParameterList()
#Set CreateWordBookmarks to true
parames.CreateWordBookmarks = True
#Create bookmarks using Headings
#parames.CreateWordBookmarksUsingHeadings = True
#Create bookmarks using word bookmarks
parames.CreateWordBookmarksUsingHeadings = False
document.SaveToFile(outputFile, parames)
document.Close()
```

---

# Spire.Doc PDF conversion with password
## Convert Word document to password-protected PDF
```python
#create word document
document = Document()

#create a parameter
toPdf = ToPdfParameterList()

#set the password
password = "E-iceblue"

toPdf.PdfSecurity.Encrypt("password", password, PdfPermissionsFlags.Default, PdfEncryptionKeySize.Key128Bit)

#save doc file.
document.SaveToFile(toPdf)
document.Close()
```

---

# Spire.Doc Python Font Color
## Change font color in Word document paragraphs
```python
#Get the first section and first paragraph
section = doc.Sections[0]
p1 = section.Paragraphs[0]

#Iterate through the childObjects of the paragraph 1 
for i in range(p1.ChildObjects.Count):
    childObj = p1.ChildObjects.get_Item(i)
    if isinstance(childObj, TextRange):
        #Change text color
        tr = childObj if isinstance(childObj, TextRange) else None
        tr.CharacterFormat.TextColor = Color.get_RosyBrown()

#Get the second paragraph
p2 = section.Paragraphs[1]

#Iterate through the childObjects of the paragraph 2
for i in range(p2.ChildObjects.Count):
    childObj = p2.ChildObjects.get_Item(i)
    if isinstance(childObj, TextRange):
        #Change text color
        tr = childObj if isinstance(childObj, TextRange) else None
        tr.CharacterFormat.TextColor = Color.get_DarkGreen()
```

---

# Spire.Doc Python Font Embedding
## Demonstrates how to embed private fonts in a Word document using Spire.Doc for Python
```python
#Get the first section and add a paragraph
section = doc.Sections[0]
p = section.AddParagraph()

#Append text to the paragraph, then set the font name and font size
txtRange = p.AppendText("Spire.Doc for Python is a professional Word Python API specifically designed for developers to create, read, write, convert, and compare Word documents with fast and high-quality performance.")
txtRange.CharacterFormat.FontName = "PT Serif Caption"
txtRange.CharacterFormat.FontSize = 20

#Allow embedding font in document
doc.EmbedFontsInFile = True

#Embed private font from font file into the document
doc.PrivateFontList.append(PrivateFontPath("PT Serif Caption", "./Data/PT Serif Caption.ttf"))
```

---

# Spire.Doc Python Font Extraction
## Extract font information from a Word document
```python
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

# Extract font information from document
font_infos = []
fontImformations = ''

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
```

---

# spire.doc python font
## set font for document text
```python
#Create a characterFormat object
characterFormat = CharacterFormat(doc)
#Set font
characterFormat.FontName = "Arial"
characterFormat.FontSize = 16

#Loop through the childObjects of paragraph 
for i in range(p.ChildObjects.Count):
    childObj = p.ChildObjects.get_Item(i)
    if isinstance(childObj, TextRange):
        #Apply character format
        tr = childObj if isinstance(childObj, TextRange) else None
        tr.ApplyCharacterFormat(characterFormat)
```

---

# spire.doc python bullet style
## create ASCII characters bullet style
```python
#Create a new document
document = Document()
section = document.AddSection()

#Create four list styles based on different ASCII characters
listStyle1 = ListStyle(document, ListType.Bulleted)
listStyle1.Name = "liststyle"
listStyle1.Levels[0].BulletCharacter = "\u006e"
listStyle1.Levels[0].CharacterFormat.FontName = "Wingdings"
document.ListStyles.Add(listStyle1)
listStyle2 = ListStyle(document, ListType.Bulleted)
listStyle2.Name = "liststyle2"
listStyle2.Levels[0].BulletCharacter = "\u0075"
listStyle2.Levels[0].CharacterFormat.FontName = "Wingdings"
document.ListStyles.Add(listStyle2)
listStyle3 = ListStyle(document, ListType.Bulleted)
listStyle3.Name = "liststyle3"
listStyle3.Levels[0].BulletCharacter = "\u00b2"
listStyle3.Levels[0].CharacterFormat.FontName = "Wingdings"
document.ListStyles.Add(listStyle3)
listStyle4 = ListStyle(document, ListType.Bulleted)
listStyle4.Name = "liststyle4"
listStyle4.Levels[0].BulletCharacter = "\u00d8"
listStyle4.Levels[0].CharacterFormat.FontName = "Wingdings"
document.ListStyles.Add(listStyle4)

#Add four paragraphs and apply list style separately
p1 = section.Body.AddParagraph()
p1.AppendText("Spire.Doc for .NET")
p1.ListFormat.ApplyStyle(listStyle1.Name)
p2 = section.Body.AddParagraph()
p2.AppendText("Spire.Doc for Java")
p2.ListFormat.ApplyStyle(listStyle2.Name)
p3 = section.Body.AddParagraph()
p3.AppendText("Spire.Doc for C++")
p3.ListFormat.ApplyStyle(listStyle3.Name)
p4 = section.Body.AddParagraph()
p4.AppendText("Spire.Doc for Python")
p4.ListFormat.ApplyStyle(listStyle4.Name)
```

---

# spire.doc python character formatting
## Apply various character formatting options to text in a Word document
```python
# Initialize document structure
document = Document()
sec = document.AddSection()
titleParagraph = sec.AddParagraph()
titleParagraph.AppendText("Font Styles and Effects ")
titleParagraph.ApplyStyle(BuiltinStyle.Title)

paragraph = sec.AddParagraph()

# Apply strikethrough formatting
tr = paragraph.AppendText("Strikethough Text")
tr.CharacterFormat.IsStrikeout = True

paragraph.AppendBreak(BreakType.LineBreak)

# Apply shadow formatting
tr = paragraph.AppendText("Shadow Text")
tr.CharacterFormat.IsShadow = True

paragraph.AppendBreak(BreakType.LineBreak)

# Apply small caps formatting
tr = paragraph.AppendText("Small caps Text")
tr.CharacterFormat.IsSmallCaps = True

paragraph.AppendBreak(BreakType.LineBreak)

# Apply double strikethrough formatting
tr = paragraph.AppendText("Double Strikethough Text")
tr.CharacterFormat.DoubleStrike = True

paragraph.AppendBreak(BreakType.LineBreak)

# Apply outline formatting
tr = paragraph.AppendText("Outline Text")
tr.CharacterFormat.IsOutLine = True

paragraph.AppendBreak(BreakType.LineBreak)

# Apply all caps formatting
tr = paragraph.AppendText("AllCaps Text")
tr.CharacterFormat.AllCaps = True

paragraph.AppendBreak(BreakType.LineBreak)

# Apply subscript formatting
tr = paragraph.AppendText("Text")
tr = paragraph.AppendText("SubScript")
tr.CharacterFormat.SubSuperScript = SubSuperScript.SubScript

tr = paragraph.AppendText("And")
tr = paragraph.AppendText("SuperScript")
tr.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript

paragraph.AppendBreak(BreakType.LineBreak)

# Apply emboss formatting with white text color
tr = paragraph.AppendText("Emboss Text")
tr.CharacterFormat.Emboss = True
tr.CharacterFormat.TextColor = Color.get_White()

paragraph.AppendBreak(BreakType.LineBreak)

# Apply hidden formatting
tr = paragraph.AppendText("Hidden:")
tr = paragraph.AppendText("Hidden Text")
tr.CharacterFormat.Hidden = True

paragraph.AppendBreak(BreakType.LineBreak)

# Apply engrave formatting with white text color
tr = paragraph.AppendText("Engrave Text")
tr.CharacterFormat.Engrave = True
tr.CharacterFormat.TextColor = Color.get_White()

paragraph.AppendBreak(BreakType.LineBreak)

# Apply font settings for Western and Chinese characters
tr = paragraph.AppendText("WesternFonts中文字体")
tr.CharacterFormat.FontNameAscii = "Calibri"
tr.CharacterFormat.FontNameNonFarEast = "Calibri"
tr.CharacterFormat.FontNameFarEast = "Simsun"

paragraph.AppendBreak(BreakType.LineBreak)

# Apply font size
tr = paragraph.AppendText("Font Size")
tr.CharacterFormat.FontSize = 20

paragraph.AppendBreak(BreakType.LineBreak)

# Apply font color
tr = paragraph.AppendText("Font Color")
tr.CharacterFormat.TextColor = Color.get_Red()

paragraph.AppendBreak(BreakType.LineBreak)

# Apply bold and italic formatting
tr = paragraph.AppendText("Bold Italic Text")
tr.CharacterFormat.Bold = True
tr.CharacterFormat.Italic = True

paragraph.AppendBreak(BreakType.LineBreak)

# Apply underline style
tr = paragraph.AppendText("Underline Style")
tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single

paragraph.AppendBreak(BreakType.LineBreak)

# Apply highlight color
tr = paragraph.AppendText("Highlight Text")
tr.CharacterFormat.HighlightColor = Color.get_Yellow()

paragraph.AppendBreak(BreakType.LineBreak)

# Apply text background color
tr = paragraph.AppendText("Text has shading")
tr.CharacterFormat.TextBackgroundColor = Color.get_Green()

paragraph.AppendBreak(BreakType.LineBreak)

# Apply text border
tr = paragraph.AppendText("Border Around Text")
tr.CharacterFormat.Border.BorderType = BorderStyle.Single

paragraph.AppendBreak(BreakType.LineBreak)

# Apply text scale
tr = paragraph.AppendText("Text Scale")
tr.CharacterFormat.TextScale = 150

paragraph.AppendBreak(BreakType.LineBreak)

# Apply character spacing
tr = paragraph.AppendText("Character Spacing is 2 point")
tr.CharacterFormat.CharacterSpacing = 2
```

---

# Spire.Doc Python Style Copy
## Copy styles from one document to another
```python
#Get the style collections of source document
styles = srcDoc.Styles

#Add the style to destination document
for i in range(styles.Count):
    style = styles.get_Item(i)
    destDoc.Styles.Add(style)
```

---

# Spire.Doc Python Character Spacing
## Extract character spacing information from a Word document
```python
# Create a document
document = Document()

# Get the first section of document
section = document.Sections[0]

# Get the first paragraph 
paragraph = section.Paragraphs[0]

# Define two variables
fontName = ""
fontSpacing = 0

# Traverse the ChildObjects 
for i in range(paragraph.ChildObjects.Count):
    docObj = paragraph.ChildObjects.get_Item(i)
    # If it is TextRange
    if isinstance(docObj, TextRange):
        textRange = docObj if isinstance(docObj, TextRange) else None

        # Get the font name
        fontName = textRange.CharacterFormat.FontName

        # Get the character spacing
        fontSpacing = textRange.CharacterFormat.CharacterSpacing
```

---

# spire.doc python text extraction
## extract text by style name from document
```python
#Create string builder
builder = ""

#Loop through sections
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    #Loop through paragraphs
    for j in range(section.Paragraphs.Count):
        para = section.Paragraphs.get_Item(j)
        #Find the paragraph whose style name is "Heading1"
        if para.StyleName == "Heading1":
            #Write the text of paragraph
            builder += para.Text
            builder += "\n"
```

---

# Spire.Doc Python Lists
## Create numbered and bulleted lists in Word documents
```python
#Initialize a document
document = Document()
#Add a section
sec = document.AddSection()
#Add paragraph and set list style
paragraph = sec.AddParagraph()
paragraph.AppendText("Lists")
paragraph.ApplyStyle(BuiltinStyle.Title)

paragraph = sec.AddParagraph()
paragraph.AppendText("Numbered List:").CharacterFormat.Bold = True

#Create list style
numberList = ListStyle(document, ListType.Numbered)
numberList.Name = "numberList"
#%1-%9
numberList.Levels[1].NumberPrefix = "%1."
numberList.Levels[1].PatternType = ListPatternType.Arabic
numberList.Levels[2].NumberPrefix = "%1.%2."
numberList.Levels[2].PatternType = ListPatternType.Arabic

bulletList = ListStyle(document, ListType.Bulleted)
bulletList.Name = "bulletList"

#add the list style into document
document.ListStyles.Add(numberList)
document.ListStyles.Add(bulletList)

#Add paragraph and apply the list style
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 1")
paragraph.ListFormat.ApplyStyle(numberList.Name)

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2")
paragraph.ListFormat.ApplyStyle(numberList.Name)

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.1")
paragraph.ListFormat.ApplyStyle(numberList.Name)
paragraph.ListFormat.ListLevelNumber = 1

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.2")
paragraph.ListFormat.ApplyStyle(numberList.Name)
paragraph.ListFormat.ListLevelNumber = 1

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.2.1")
paragraph.ListFormat.ApplyStyle(numberList.Name)
paragraph.ListFormat.ListLevelNumber = 2
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.2.2")
paragraph.ListFormat.ApplyStyle(numberList.Name)
paragraph.ListFormat.ListLevelNumber = 2
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.2.3")
paragraph.ListFormat.ApplyStyle(numberList.Name)
paragraph.ListFormat.ListLevelNumber = 2

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.3")
paragraph.ListFormat.ApplyStyle(numberList.Name)
paragraph.ListFormat.ListLevelNumber = 1

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 3")
paragraph.ListFormat.ApplyStyle(numberList.Name)

paragraph = sec.AddParagraph()
paragraph.AppendText("Bulleted List:").CharacterFormat.Bold = True

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 1")
paragraph.ListFormat.ApplyStyle(bulletList.Name)
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2")
paragraph.ListFormat.ApplyStyle(bulletList.Name)

paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.1")
paragraph.ListFormat.ApplyStyle(bulletList.Name)
paragraph.ListFormat.ListLevelNumber = 1
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 2.2")
paragraph.ListFormat.ApplyStyle(bulletList.Name)
paragraph.ListFormat.ListLevelNumber = 1
paragraph = sec.AddParagraph()
paragraph.AppendText("List Item 3")
paragraph.ListFormat.ApplyStyle(bulletList.Name)
```

---

# Spire.Doc Python styles
## Apply multiple styles within a paragraph
```python
# Create a Word document
doc = Document()

# Add a section
section = doc.AddSection()

# Add a paragraph
para = section.AddParagraph()

# Add a text range 1 and set its style
range = para.AppendText("Spire.Doc for .NET ")
range.CharacterFormat.FontName = "Calibri"
range.CharacterFormat.FontSize = 16
range.CharacterFormat.TextColor = Color.get_Blue()
range.CharacterFormat.Bold = True
range.CharacterFormat.UnderlineStyle = UnderlineStyle.Single

# Add a text range 2 and set its style
range = para.AppendText("is a professional Word .NET library")
range.CharacterFormat.FontName = "Calibri"
range.CharacterFormat.FontSize = 15
```

---

# Spire.Doc Python Paragraph Formatting
## Demonstrates various paragraph formatting options in a Word document
```python
#Initialize a document
document = Document()
sec = document.AddSection()
para = sec.AddParagraph()
para.AppendText("Paragraph Formatting")
para.ApplyStyle(BuiltinStyle.Title)

para = sec.AddParagraph()
para.AppendText("This paragraph is surrounded with borders.")
para.Format.Borders.BorderType = BorderStyle.Single
para.Format.Borders.Color = Color.get_Red()

para = sec.AddParagraph()
para.AppendText("The alignment of this paragraph is Left.")
para.Format.HorizontalAlignment = HorizontalAlignment.Left

para = sec.AddParagraph()
para.AppendText("The alignment of this paragraph is Center.")
para.Format.HorizontalAlignment = HorizontalAlignment.Center

para = sec.AddParagraph()
para.AppendText("The alignment of this paragraph is Right.")
para.Format.HorizontalAlignment = HorizontalAlignment.Right

para = sec.AddParagraph()
para.AppendText("The alignment of this paragraph is justified.")
para.Format.HorizontalAlignment = HorizontalAlignment.Justify

para = sec.AddParagraph()
para.AppendText("The alignment of this paragraph is distributed.")
para.Format.HorizontalAlignment = HorizontalAlignment.Distribute

para = sec.AddParagraph()
para.AppendText("This paragraph has the gray shadow.")
para.Format.BackColor = Color.get_Gray()

para = sec.AddParagraph()
para.AppendText("This paragraph has the following indentations: Left indentation is 10pt, right indentation is 10pt, first line indentation is 15pt.")
para.Format.SetLeftIndent(10)
para.Format.SetRightIndent(10)
para.Format.SetFirstLineIndent(15)

para = sec.AddParagraph()
para.AppendText("The hanging indentation of this paragraph is 15pt.")
#Negative value represents hanging indentation
para.Format.SetFirstLineIndent(-15)

para = sec.AddParagraph()
para.AppendText("This paragraph has the following spacing: spacing before is 10pt, spacing after is 20pt, line spacing is at least 10pt.")
para.Format.AfterSpacing = 20
para.Format.BeforeSpacing = 10
para.Format.LineSpacingRule = LineSpacingRule.AtLeast
para.Format.LineSpacing = 10
```

---

# Spire.Doc Python Restart List Numbering
## Create numbered lists with restart functionality
```python
#Create word document
document = Document()

#Create a new section
section = document.AddSection()

#Create a new paragraph
paragraph = section.AddParagraph()

#Append Text
paragraph.AppendText("List 1")

numberList = ListStyle(document, ListType.Numbered)
numberList.Name = "Numbered1"
document.ListStyles.Add(numberList)

#Add paragraph and apply the list style
paragraph = section.AddParagraph()
paragraph.AppendText("List Item 1")
paragraph.ListFormat.ApplyStyle(numberList.Name)

paragraph = section.AddParagraph()
paragraph.AppendText("List Item 2")
paragraph.ListFormat.ApplyStyle(numberList.Name)

paragraph = section.AddParagraph()
paragraph.AppendText("List Item 3")
paragraph.ListFormat.ApplyStyle(numberList.Name)

paragraph = section.AddParagraph()
paragraph.AppendText("List Item 4")
paragraph.ListFormat.ApplyStyle(numberList.Name)

#Append Text
paragraph = section.AddParagraph()
paragraph.AppendText("List 2")

numberList2 = ListStyle(document, ListType.Numbered)
numberList2.Name = "Numbered2"
#set start number of second list
numberList2.Levels[0].StartAt = 10
document.ListStyles.Add(numberList2)

#Add paragraph and apply the list style
paragraph = section.AddParagraph()
paragraph.AppendText("List Item 5")
paragraph.ListFormat.ApplyStyle(numberList2.Name)

paragraph = section.AddParagraph()
paragraph.AppendText("List Item 6")
paragraph.ListFormat.ApplyStyle(numberList2.Name)

paragraph = section.AddParagraph()
paragraph.AppendText("List Item 7")
paragraph.ListFormat.ApplyStyle(numberList2.Name)

paragraph = section.AddParagraph()
paragraph.AppendText("List Item 8")
paragraph.ListFormat.ApplyStyle(numberList2.Name)
```

---

# Spire.Doc Python Style Retrieval
## Extract style names from paragraphs in a document
```python
# Initialize a document
doc = Document()

# Traverse all paragraphs in the document and get their style names through StyleName property
styleName = ''
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    for j in range(section.Paragraphs.Count):
        paragraph = section.Paragraphs.get_Item(j)
        styleName += paragraph.StyleName + "\r\n"
```

---

# Spire.Doc Document Styling
## Creating and modifying document styles in Word documents

```python
# Initialize a document
document = Document()
document.AddSection()

# Title style customization
titleStyle = document.AddStyle(BuiltinStyle.Title)
titleStyle.CharacterFormat.FontName = "cambria"
titleStyle.CharacterFormat.FontSize = 28
titleStyle.CharacterFormat.TextColor = Color.FromArgb(255, 42, 123, 136)

# Paragraph format for title style
if isinstance(titleStyle, ParagraphStyle):
    ps = titleStyle if isinstance(titleStyle, ParagraphStyle) else None
    ps.ParagraphFormat.Borders.Bottom.BorderType = BorderStyle.Single
    ps.ParagraphFormat.Borders.Bottom.Color = Color.FromArgb(255, 42, 123, 136)
    ps.ParagraphFormat.Borders.Bottom.LineWidth = 1.5
    ps.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left

# Normal style customization
normalStyle = document.AddStyle(BuiltinStyle.Normal)
normalStyle.CharacterFormat.FontName = "cambria"
normalStyle.CharacterFormat.FontSize = 11

# Heading1 style customization
heading1Style = document.AddStyle(BuiltinStyle.Heading1)
heading1Style.CharacterFormat.FontName = "cambria"
heading1Style.CharacterFormat.FontSize = 14
heading1Style.CharacterFormat.Bold = True
heading1Style.CharacterFormat.TextColor = Color.FromArgb(255, 42, 123, 136)

# Heading2 style customization
heading2Style = document.AddStyle(BuiltinStyle.Heading2)
heading2Style.CharacterFormat.FontName = "cambria"
heading2Style.CharacterFormat.FontSize = 12
heading2Style.CharacterFormat.Bold = True

# Custom list style creation
bulletList = ListStyle(document, ListType.Bulleted)
bulletList.CharacterFormat.FontName = "cambria"
bulletList.CharacterFormat.FontSize = 12
bulletList.Name = "bulletList"
document.ListStyles.Add(bulletList)
```

---

# spire.doc mail merge locale change
## change locale settings for mail merge operations
```python
# Store the current culture so it can be set back once mail merge is complete.
current_locale = locale.getlocale()

locale.setlocale(locale.LC_ALL,'de_DE.UTF-8')

fieldNames = ["Contact Name", "Fax", "Date"]
fieldValues = ["John Smith", "+1 (69) 123456", datetime.datetime.now().strftime('%c')]
document.MailMerge.Execute(fieldNames, fieldValues)

locale.setlocale(locale.LC_ALL,current_locale)
```

---

# Spire.Doc Python Conditional Fields
## Create and execute conditional IF fields in a document using mail merge
```python
def _CreateIFField1(document, paragraph):
    ifField = IfField(document)
    ifField.Type = FieldType.FieldIf
    ifField.Code = "IF "
    paragraph.Items.Add(ifField)

    paragraph.AppendField("Count", FieldType.FieldMergeField)
    paragraph.AppendText(" > ")
    paragraph.AppendText("\"1\" ")
    paragraph.AppendText("\"Greater than one\" ")
    paragraph.AppendText("\"Less than one\"")

    end = document.CreateParagraphItem(ParagraphItemType.FieldMark)
    tempFieldMark = ( end if isinstance(end, FieldMark) else None)
    if tempFieldMark != None:
        tempFieldMark.Type = FieldMarkType.FieldEnd
        
    paragraph.Items.Add(end)
    ifField.End = end if isinstance(end, FieldMark) else None

def _CreateIFField2(document, paragraph):
    ifField = IfField(document)
    ifField.Type = FieldType.FieldIf
    ifField.Code = "IF "
    paragraph.Items.Add(ifField)

    paragraph.AppendField("Age", FieldType.FieldMergeField)
    paragraph.AppendText(" > ")
    paragraph.AppendText("\"50\" ")
    paragraph.AppendText("\"The old man\" ")
    paragraph.AppendText("\"The young man\"")

    end = document.CreateParagraphItem(ParagraphItemType.FieldMark)
    tempFieldMark = ( end if isinstance(end, FieldMark) else None)
    tempFieldMark.Type = FieldMarkType.FieldEnd
    paragraph.Items.Add(end)

    ifField.End = end if isinstance(end, FieldMark) else None

# Create document and add conditional fields
doc = Document()
section = doc.AddSection()
paragraph = section.AddParagraph()

_CreateIFField1(doc, paragraph)
paragraph = section.AddParagraph()
_CreateIFField2(doc, paragraph)

# Execute mail merge with conditional fields
fieldName = ["Count", "Age"]
fieldValue = ["2", "30"]

doc.MailMerge.Execute(fieldName, fieldValue)
doc.IsUpdateFields = True
```

---

# Spire.Doc Python Mail Merge
## Hide empty regions in mail merge
```python
#Set the value to remove paragraphs which contain empty field.
document.MailMerge.HideEmptyParagraphs = True
#Set the value to remove group which contain empty field.
document.MailMerge.HideEmptyGroup = True
document.MailMerge.Execute(filedNames, filedValues)
```

---

# spire.doc python mail merge
## identify merge field names in a Word document
```python
#Get the collection of group names.
GroupNames = document.MailMerge.GetMergeGroupNames()

#Get the collection of merge field names in a specific group.
MergeFieldNamesWithinRegion = document.MailMerge.GetMergeFieldNames("Products")

#Get the collection of all the merge field names.
MergeFieldNames = document.MailMerge.GetMergeFieldNames()
```

---

# Spire.Doc Python Mail Merge
## Execute mail merge operation in a Word document
```python
# Execute mail merge with field names and values
document.MailMerge.Execute(filedNames, filedValues)
```

---

# Spire.Doc Python Mail Merge
## Execute mail merge with field names and values
```python
doc = Document()
fieldName = ["XX_Name"]
fieldValue = ["Jason Tang"]
doc.MailMerge.Execute(fieldName, fieldValue)
```

---

# Spire.Doc Python Nested Mail Merge
## Execute nested mail merge operation with XML data source
```python
# execute mailmerge
tempdDict = {"Customer": '', "Order": "Customer_Id = %Customer.Customer_Id%"}
dataFile = "Data/Orders.xml"
document.MailMerge.ExecuteWidthNestedRegion(dataFile, tempdDict)
```

---

# Copy Bookmark Content
## This code demonstrates how to copy content from a bookmark in a Word document.
```python
# Get the bookmark by name.
bookmark = doc.Bookmarks["Test"]
docObj = None

# Judge if the paragraph includes the bookmark exists in the table, if it exists in cell,
# then need to find its outermost parent object(Table),
# and get the start/end index of current object on body.
if ( bookmark.BookmarkStart.Owner if isinstance(bookmark.BookmarkStart.Owner, Paragraph) else None).IsInCell:
    docObj = bookmark.BookmarkStart.Owner.Owner.Owner.Owner
else:
    docObj = bookmark.BookmarkStart.Owner
startIndex = doc.Sections[0].Body.ChildObjects.IndexOf(docObj)
if ( bookmark.BookmarkEnd.Owner if isinstance(bookmark.BookmarkEnd.Owner, Paragraph) else None).IsInCell:
    docObj = bookmark.BookmarkEnd.Owner.Owner.Owner.Owner
else:
    docObj = bookmark.BookmarkEnd.Owner
endIndex = doc.Sections[0].Body.ChildObjects.IndexOf(docObj)

# Get the start/end index of the bookmark object on the paragraph.
para = bookmark.BookmarkStart.Owner if isinstance(bookmark.BookmarkStart.Owner, Paragraph) else None
pStartIndex = para.ChildObjects.IndexOf(bookmark.BookmarkStart)
para = bookmark.BookmarkEnd.Owner if isinstance(bookmark.BookmarkEnd.Owner, Paragraph) else None
pEndIndex = para.ChildObjects.IndexOf(bookmark.BookmarkEnd)

# Get the content of current bookmark and copy.
select = TextBodySelection(doc.Sections[0].Body, startIndex, endIndex, pStartIndex, pEndIndex)
body = TextBodyPart(select)
for i in range(body.BodyItems.Count):
    doc.Sections[0].Body.ChildObjects.Add(body.BodyItems[i].Clone())
```

---

# Spire.Doc Python Bookmark Creation
## Create simple and nested bookmarks in Word documents
```python
def _CreateBookmark(section):
    paragraph = section.AddParagraph()
    txtRange = paragraph.AppendText("The following example demonstrates how to create bookmark in a Word document.")
    txtRange.CharacterFormat.Italic = True

    section.AddParagraph()
    paragraph = section.AddParagraph()
    txtRange = paragraph.AppendText("Simple Create Bookmark.")
    txtRange.CharacterFormat.TextColor = Color.get_CornflowerBlue()
    paragraph.ApplyStyle(BuiltinStyle.Heading2)

    #Write simple CreateBookmarks.
    section.AddParagraph()
    paragraph = section.AddParagraph()
    paragraph.AppendBookmarkStart("SimpleCreateBookmark")
    paragraph.AppendText("This is a simple bookmark.")
    paragraph.AppendBookmarkEnd("SimpleCreateBookmark")

    section.AddParagraph()
    paragraph = section.AddParagraph()
    txtRange = paragraph.AppendText("Nested Create Bookmark.")
    txtRange.CharacterFormat.TextColor = Color.get_CornflowerBlue()
    paragraph.ApplyStyle(BuiltinStyle.Heading2)

    #Write nested CreateBookmarks.
    section.AddParagraph()
    paragraph = section.AddParagraph()
    paragraph.AppendBookmarkStart("Root")
    txtRange = paragraph.AppendText(" This is Root data ")
    txtRange.CharacterFormat.Italic = True
    paragraph.AppendBookmarkStart("NestedLevel1")
    txtRange = paragraph.AppendText(" This is Nested Level1 ")
    txtRange.CharacterFormat.Italic = True
    txtRange.CharacterFormat.TextColor = Color.get_DarkSlateGray()
    paragraph.AppendBookmarkStart("NestedLevel2")
    txtRange = paragraph.AppendText(" This is Nested Level2 ")
    txtRange.CharacterFormat.Italic = True
    txtRange.CharacterFormat.TextColor = Color.get_DimGray()
    paragraph.AppendBookmarkEnd("NestedLevel2")
    paragraph.AppendBookmarkEnd("NestedLevel1")
    paragraph.AppendBookmarkEnd("Root")
```

---

# Spire.Doc Python Bookmark
## Create bookmark for table in Word document
```python
def _CreateBookmarkForTable(doc, section):
    #Add a paragraph
    paragraph = section.AddParagraph()

    #Append text for added paragraph
    txtRange = paragraph.AppendText("The following example demonstrates how to create bookmark for a table in a Word document.")

    #Set the font in italic
    txtRange.CharacterFormat.Italic = True

    #Append bookmark start
    paragraph.AppendBookmarkStart("CreateBookmark")

    #Append bookmark end
    paragraph.AppendBookmarkEnd("CreateBookmark")

    #Add table
    table = section.AddTable(True)

    #Set the number of rows and columns
    table.ResetCells(2, 2)

    #Append text for table cells
    range = table.Rows.get_Item(0).Cells.get_Item(0).AddParagraph().AppendText("sampleA")
    range = table.Rows.get_Item(0).Cells.get_Item(1).AddParagraph().AppendText("sampleB")
    range = table.Rows.get_Item(1).Cells.get_Item(0).AddParagraph().AppendText("120")
    range = table.Rows.get_Item(1).Cells.get_Item(1).AddParagraph().AppendText("260")

    #Get the bookmark by index.
    bookmark = doc.Bookmarks[0]

    #Get the name of bookmark.
    bookmarkName = bookmark.Name

    #Locate the bookmark by name.
    navigator = BookmarksNavigator(doc)
    navigator.MoveToBookmark(bookmarkName)

    #Add table to TextBodyPart
    part = navigator.GetBookmarkContent()
    part.BodyItems.Add(table)

    #Replace bookmark cotent with table
    navigator.ReplaceBookmarkContent(part)
```

---

# spire.doc python bookmark text extraction
## Extract text content from a specific bookmark in a document
```python
#Creates a BookmarkNavigator instance to access the bookmark
navigator = BookmarksNavigator(doc)
#Locate a specific bookmark by bookmark name
navigator.MoveToBookmark("Content")
textBodyPart = navigator.GetBookmarkContent()

#Iterate through the items in the bookmark content to get the text
text = ''
for i in range(textBodyPart.BodyItems.Count):
    item = textBodyPart.BodyItems.get_Item(i)
    if isinstance(item, Paragraph):
        for j in range(( item if isinstance(item, Paragraph) else None).ChildObjects.Count):
            childObject = ( item if isinstance(item, Paragraph) else None).ChildObjects.get_Item(j)
            if isinstance(childObject, TextRange):
                text += ( childObject if isinstance(childObject, TextRange) else None).Text
```

---

# spire.doc python bookmarks
## get bookmarks from word document
```python
#Get the bookmark by index.
bookmark1 = document.Bookmarks[0]

#Get the bookmark by name.
bookmark2 = document.Bookmarks["Test2"]
```

---

# spire.doc insert document at bookmark
## insert content from one document at a bookmark location in another document
```python
#Create the first document
document1 = Document()

#Create the second document
document2 = Document()

#Get the first section of the first document 
section1 = document1.Sections[0]

#Locate the bookmark
bn = BookmarksNavigator(document1)

#Find bookmark by name
bn.MoveToBookmark("Test", True, True)

#Get bookmarkStart
start = bn.CurrentBookmark.BookmarkStart

#Get the owner paragraph
para = start.OwnerParagraph

#Get the para index
index = section1.Body.ChildObjects.IndexOf(para)

#Insert the paragraphs of document2
for i in range(document2.Sections.Count):
    section2 = document2.Sections.get_Item(i)
    for j in range(section2.Paragraphs.Count):
        paragraph = section2.Paragraphs.get_Item(j)
        cloneP = paragraph.Clone()
        section1.Body.ChildObjects.Insert(index + 1, cloneP if isinstance(cloneP, Paragraph) else None)
```

---

# Spire.Doc Python Bookmark Image Insertion
## Insert an image at a specific bookmark location in a Word document
```python
# Create a document
doc = Document()

# Create an instance of BookmarksNavigator
bn = BookmarksNavigator(doc)

# Find a bookmark named Test
bn.MoveToBookmark("Test", True, True)

# Add a section
section0 = doc.AddSection()

# Add a paragraph for the section
paragraph = section0.AddParagraph()

# Add a picture into the paragraph
picture = paragraph.AppendPicture("./Data/Word.png")

# Add the paragraph at the position of bookmark
bn.InsertParagraph(paragraph)

# Remove the section0
doc.Sections.Remove(section0)
```

---

# Spire.Doc Python Bookmark Removal
## Remove a bookmark from a Word document
```python
#Create a document
document = Document()

#Get the bookmark by name.
bookmark = document.Bookmarks["Test"]

#Remove the bookmark, not its content.
document.Bookmarks.Remove(bookmark)
```

---

# Spire.Doc Python Bookmark Operations
## Remove content between bookmark start and end
```python
#Get the bookmark by name.            
bookmark = document.Bookmarks["Test"]

para = bookmark.BookmarkStart.Owner if isinstance(bookmark.BookmarkStart.Owner, Paragraph) else None
startIndex = para.ChildObjects.IndexOf(bookmark.BookmarkStart)
para = bookmark.BookmarkEnd.Owner if isinstance(bookmark.BookmarkEnd.Owner, Paragraph) else None
endIndex = para.ChildObjects.IndexOf(bookmark.BookmarkEnd)

#Remove the content object, and Start from next of BookmarkStart object, end up with previous of BookmarkEnd object. 
#This method is only to remove the content of the bookmark.
for i in range(startIndex + 1, endIndex):
    para.ChildObjects.RemoveAt(startIndex + 1)
```

---

# spire.doc python bookmark
## replace bookmark content
```python
#Locate the bookmark.
bookmarkNavigator = BookmarksNavigator(doc)
bookmarkNavigator.MoveToBookmark("Test")

#Replace the context with new.
bookmarkNavigator.ReplaceBookmarkContent("This is replaced content.", False)
```

---

# Spire.Doc Python Bookmark Replacement
## Replace bookmark content with a table
```python
#Create a table
table = Table(doc, True)

#Get the specific bookmark by its name
navigator = BookmarksNavigator(doc)
navigator.MoveToBookmark("Test")

#Create a TextBodyPart instance and add the table to it
part = TextBodyPart(doc)
part.BodyItems.Add(table)

#Replace the current bookmark content with the TextBodyPart object
navigator.ReplaceBookmarkContent(part)
```

---

# Spire.Doc Comments
## Add comments for specific text in a document
```python
def InsertComments(doc, keystring):
    #Find the key string
    find = doc.FindString(keystring, False, True)

    #Create the commentmarkStart and commentmarkEnd
    commentmarkStart = CommentMark(doc)
    commentmarkStart.Type = CommentMarkType.CommentStart
    commentmarkStart.CommentId = 1
    commentmarkEnd = CommentMark(doc)
    commentmarkEnd.CommentId = 1
    commentmarkEnd.Type = CommentMarkType.CommentEnd

    #Add the content for comment
    comment = Comment(doc)
    comment.Format.CommentId = 1
    comment.Body.AddParagraph().Text = "Test comments"
    comment.Format.Author = "E-iceblue"

    #Get the textRanges
    text_ranges = find.GetRanges()
    length = len(text_ranges)
 
    #Get its paragraph
    para = text_ranges[0].OwnerParagraph

    #Get the index of textRange 
    index = para.ChildObjects.IndexOf(text_ranges[0])
    
    #Insert the commentmarkStart and commentmarkEnd
    para.ChildObjects.Insert(index, commentmarkStart)
    para.ChildObjects.Insert(index + length+1, commentmarkEnd)
    para.ChildObjects.Add(comment)
```

---

# spire.doc python comment
## insert comment into word document
```python
# Insert comment.
paragraph = section.Paragraphs[1]
comment = paragraph.AppendComment("Spire.Doc for .NET")
comment.Format.Author = "E-iceblue"
comment.Format.Initial = "CM"
```

---

# Extract comments from document
## This code demonstrates how to extract all comments from a Word document
```python
doc = Document()

content = ''

#Traverse all comments
for i in range(doc.Comments.Count):
    comment = doc.Comments.get_Item(i)
    for j in range(comment.Body.Paragraphs.Count):
        p = comment.Body.Paragraphs.get_Item(j)
        content += p.Text
        content += '\n'
```

---

# Insert Picture into Comment
## This code demonstrates how to insert a picture into a comment in a Word document
```python
#Get the first paragraph and insert comment
paragraph = doc.Sections[0].Paragraphs[2]
comment = paragraph.AppendComment("This is a comment.")
comment.Format.Author = "E-iceblue"

#Load a picture
docPicture = DocPicture(doc)
docPicture.LoadImage("./Data/E-iceblue.png")
#Insert the picture into the comment body
comment.Body.AddParagraph().ChildObjects.Add(docPicture)
```

---

# Spire.Doc Python Comment Management
## Core functionality for removing and replacing comments in Word documents
```python
#Replace the content of the first comment
doc.Comments[0].Body.Paragraphs[0].Replace("This is the title", "This comment is changed.", False, False)

#Remove the second comment
doc.Comments.RemoveAt(1)
```

---

# Spire.Doc Python Comment Processing
## Remove content associated with comments in a Word document
```python
#Get the first comment
comment = document.Comments[0]

#Get the paragraph of obtained comment
para = comment.OwnerParagraph

#Get index of the CommentMarkStart 
startIndex = para.ChildObjects.IndexOf(comment.CommentMarkStart)

#Get index of the CommentMarkEnd
endIndex = para.ChildObjects.IndexOf(comment.CommentMarkEnd)

#Create a list
dataList = []

#Get TextRanges between the indexes
for i in range(startIndex, endIndex):
    if isinstance(para.ChildObjects[i], TextRange):
        dataList.append(para.ChildObjects[i] if isinstance(para.ChildObjects[i], TextRange) else None)

#Insert a new TextRange
textRange = TextRange(document)

#Set text is null
textRange.Text = None

#Insert the new textRange
para.ChildObjects.Insert(endIndex, textRange)

#Remove previous TextRanges
for i, unusedItem in enumerate(dataList):
    para.ChildObjects.Remove(dataList[i])
```

---

# Spire.Doc Python Comment Reply
## Create and add a reply to a comment in a Word document
```python
#get the first comment.
comment1 = doc.Comments[0]

#create a new comment and specify the author and content.
replyComment1 = Comment(doc)
replyComment1.Format.Author = "E-iceblue"
replyComment1.Body.AddParagraph().AppendText("Spire.Doc is a professional Word .NET library on operating Word documents.")

#add the new comment as a reply to the selected comment.
comment1.ReplyToComment(replyComment1)

#create a picture
docPicture = DocPicture(doc)

#insert a picture in the comment
replyComment1.Body.Paragraphs[0].ChildObjects.Add(docPicture)
```

---

# spire.doc python barcode
## add barcode image to word document
```python
#Add barcode image
picture = document.Sections[0].AddParagraph().AppendPicture(imgPath)
```

---

# Spire.Doc Python Horizontal Line
## Add a horizontal line to a Word document
```python
# Create Word document
doc = Document()
sec = doc.AddSection()
para = sec.AddParagraph()
para.AppendHorizonalLine()
```

---

# Spire.Doc Python Image and Shape
## Add image and textbox to footer of Word document
```python
# Image path
imgPath = "./Data/Spire.Doc.png"

#Add a picture in footer and set it's position
picture = document.Sections[0].HeadersFooters.Footer.AddParagraph().AppendPicture(imgPath)

picture.VerticalOrigin = VerticalOrigin.Page
picture.HorizontalOrigin = HorizontalOrigin.Page
picture.VerticalAlignment = ShapeVerticalAlignment.Bottom
picture.TextWrappingStyle = TextWrappingStyle.none

#Add a textbox in footer and set it's positiion
textbox = document.Sections[0].HeadersFooters.Footer.AddParagraph().AppendTextBox(150, 20)
textbox.VerticalOrigin = VerticalOrigin.Page
textbox.HorizontalOrigin = HorizontalOrigin.Page
textbox.HorizontalPosition = 300
textbox.VerticalPosition = 700
textbox.Body.AddParagraph().AppendText("Welcome to E-iceblue")
```

---

# Spire.Doc Python Shape Group
## Add shape group with text boxes and arrows to document
```python
#create a document
doc = Document()
sec = doc.AddSection()

#add a new paragraph
para = sec.AddParagraph()
#add a shape group with the height and width
shapegroup = para.AppendShapeGroup(375, 462)
shapegroup.HorizontalPosition = 180
#calculate the scale ratio
X = float((shapegroup.Width / 1000.0))
Y = float((shapegroup.Height / 1000.0))

txtBox = TextBox(doc)
txtBox.SetShapeType(ShapeType.RoundRectangle)
txtBox.Width = 125 / X
txtBox.Height = 54 / Y
paragraph = txtBox.Body.AddParagraph()
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
paragraph.AppendText("Start")
txtBox.HorizontalPosition = 19 / X
txtBox.VerticalPosition = 27 / Y
txtBox.Format.LineColor = Color.get_Green()
shapegroup.ChildObjects.Add(txtBox)

arrowLineShape = ShapeObject(doc, ShapeType.DownArrow)
arrowLineShape.Width = 16 / X
arrowLineShape.Height = 40 / Y
arrowLineShape.HorizontalPosition = 69 / X
arrowLineShape.VerticalPosition = 87 / Y
arrowLineShape.StrokeColor = Color.get_Purple()
shapegroup.ChildObjects.Add(arrowLineShape)

txtBox = TextBox(doc)
txtBox.SetShapeType(ShapeType.Rectangle)
txtBox.Width = 125 / X
txtBox.Height = 54 / Y
paragraph = txtBox.Body.AddParagraph()
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
paragraph.AppendText("Step 1")
txtBox.HorizontalPosition = 19 / X
txtBox.VerticalPosition = 131 / Y
txtBox.Format.LineColor = Color.get_Blue()
shapegroup.ChildObjects.Add(txtBox)

arrowLineShape = ShapeObject(doc, ShapeType.DownArrow)
arrowLineShape.Width = 16 / X
arrowLineShape.Height = 40 / Y
arrowLineShape.HorizontalPosition = 69 / X
arrowLineShape.VerticalPosition = 192 / Y
arrowLineShape.StrokeColor = Color.get_Purple()
shapegroup.ChildObjects.Add(arrowLineShape)

txtBox = TextBox(doc)
txtBox.SetShapeType(ShapeType.Parallelogram)
txtBox.Width = 149 / X
txtBox.Height = 59 / Y
paragraph = txtBox.Body.AddParagraph()
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
paragraph.AppendText("Step 2")
txtBox.HorizontalPosition = 7 / X
txtBox.VerticalPosition = 236 / Y
txtBox.Format.LineColor = Color.get_BlueViolet()
shapegroup.ChildObjects.Add(txtBox)

arrowLineShape = ShapeObject(doc, ShapeType.DownArrow)
arrowLineShape.Width = 16 / X
arrowLineShape.Height = 40 / Y
arrowLineShape.HorizontalPosition = 66 / X
arrowLineShape.VerticalPosition = 300 / Y
arrowLineShape.StrokeColor = Color.get_Purple()
shapegroup.ChildObjects.Add(arrowLineShape)

txtBox = TextBox(doc)
txtBox.SetShapeType(ShapeType.Rectangle)
txtBox.Width = 125 / X
txtBox.Height = 54 / Y
paragraph = txtBox.Body.AddParagraph()
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
paragraph.AppendText("Step 3")
txtBox.HorizontalPosition = 19 / X
txtBox.VerticalPosition = 345 / Y
txtBox.Format.LineColor = Color.get_Blue()
shapegroup.ChildObjects.Add(txtBox)
```

---

# spire.doc python add shapes
## Add various shapes to a Word document and arrange them in a grid pattern
```python
#Create Word document.
doc = Document()
sec = doc.AddSection()
para = sec.AddParagraph()
x = 60
y = 40
lineCount = 0
for i in range(1, 20):
    if lineCount > 0 and math.fmod(lineCount, 8) == 0:
        para.AppendBreak(BreakType.PageBreak)
        x = 60
        y = 40
        lineCount = 0
    #Add shape and set its size and position.
    shape = para.AppendShape(50, 50, ShapeType(i))
    shape.HorizontalOrigin = HorizontalOrigin.Page
    shape.HorizontalPosition = x
    shape.VerticalOrigin = VerticalOrigin.Page
    shape.VerticalPosition = y + 50
    x = x + int(shape.Width) + 50
    if i > 0 and math.fmod(i, 5) == 0:
        y = y + int(shape.Height) + 120
        lineCount += 1
        x = 60
```

---

# spire.doc python svg
## Add SVG image to Word document
```python
#Create a Word document.
document = Document()

#Create a new section
section = document.AddSection()

#Add a new paragraph
para = section.AddParagraph()

#add a svg file to the paragraph
svgPicture = para.AppendPicture(inputFile)

#Set svg's width
svgPicture.Width = 200

#Set svg's height
svgPicture.Height = 200
```

---

# spire.doc python shape alignment
## Align shapes in a Word document by setting horizontal alignment
```python
section = doc.Sections[0]

for i in range(section.Paragraphs.Count):
    para = section.Paragraphs.get_Item(i)
    for j in range(para.ChildObjects.Count):
        obj = para.ChildObjects.get_Item(j)
        if isinstance(obj, ShapeObject):
            #Set the horizontal alignment as center
            ( obj if isinstance(obj, ShapeObject) else None).HorizontalAlignment = ShapeHorizontalAlignment.Center

            #//Set the vertical alignment as top
            #(obj as ShapeObject).VerticalAlignment = ShapeVerticalAlignment.Top
```

---

# Spire.Doc Python Image Extraction
## Extract images from Word document by traversing document elements
```python
# Initialize queue for document traversal
nodes = queue.Queue()
nodes.put(document)

# List to store embedded images
images = []

# Traverse document elements
while nodes.qsize() > 0:
    node = nodes.get()
    for i in range(node.ChildObjects.Count):
        child = node.ChildObjects.get_Item(i)
        if child.DocumentObjectType == DocumentObjectType.Picture:
            picture = child if isinstance(child, DocPicture) else None
            dataBytes = picture.ImageBytes
            images.append(dataBytes)
        elif isinstance(child, ICompositeObject):
            nodes.put(child if isinstance(child, ICompositeObject) else None)
```

---

# Spire.Doc get alternative text
## Extract alternative text from shapes in a Word document
```python
#Loop through shapes and get the AlternativeText
for i in range(document.Sections.Count):
    section = document.Sections.get_Item(i)
    for j in range(section.Paragraphs.Count):
        para = section.Paragraphs.get_Item(j)
        for k in range(para.ChildObjects.Count):
            obj = para.ChildObjects.get_Item(k)
            if isinstance(obj, ShapeObject):
                text = ( obj if isinstance(obj, ShapeObject) else None).AlternativeText
                #Append the alternative text in builder
                builder += text
                builder += '\n'
```

---

# Insert Image in Word Document
## This code demonstrates how to insert an image into a Word document using Spire.Doc for Python.
```python
# Create a document
document = Document()

# Add a section
section = document.AddSection()

# Add paragraph
paragraph = section.AddParagraph()
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Left
picture = paragraph.AppendPicture("./Data/Spire.Doc.png")

picture.Width = 100
picture.Height = 100
```

---

# spire.doc python image insertion
## insert image into document with position and size settings
```python
# Create a picture
picture = DocPicture(doc)
picture.LoadImage("./Data/Word.png")

# set image's position
picture.HorizontalPosition = 50.0
picture.VerticalPosition = 60.0

# set image's size
picture.Width = 200.0
picture.Height = 200.0

# set textWrappingStyle with image
picture.TextWrappingStyle = TextWrappingStyle.Through
# Insert the picture at the beginning of the paragraph
paragraph.ChildObjects.Insert(0, picture)
```

---

# Spire.Doc Python WordArt
## Insert WordArt into a Word document
```python
# Create Word document.
doc = Document()

# Add a paragraph.
paragraph = doc.Sections[0].AddParagraph()

# Add a shape.
shape = paragraph.AppendShape(250, 70, ShapeType.TextWave4)

# Set the position of the shape.
shape.VerticalPosition = 20
shape.HorizontalPosition = 80

# set the text of WordArt.
shape.WordArt.Text = "Thanks for reading."

# Set the fill color.
shape.FillColor = Color.get_Red()

# Set the border color of the text.
shape.StrokeColor = Color.get_Yellow()
```

---

# spire.doc python image replacement
## replace images with text in word document
```python
# Replace all pictures with texts
j = 1
for k in range(doc.Sections.Count):
    sec = doc.Sections.get_Item(k)
    for m in range(sec.Paragraphs.Count):
        para = sec.Paragraphs.get_Item(m)
        pictures = []
        # Get all pictures in the Word document
        for x in range(para.ChildObjects.Count):
            docObj = para.ChildObjects.get_Item(x)
            if docObj.DocumentObjectType == DocumentObjectType.Picture:
                pictures.append(docObj)

        # Replace pictures with the text "Here was image {image index}"
        for pic in pictures:
            index = para.ChildObjects.IndexOf(pic)
            textRange = TextRange(doc)
            textRange.Text = "Here was image {0}".format(j)
            para.ChildObjects.Insert(index, textRange)
            para.ChildObjects.Remove(pic)
            j += 1
```

---

# Spire.Doc Python Reset Image Size
## Reset the size of images in a Word document
```python
# Get the first section
section = doc.Sections[0]
# Get the first paragraph
paragraph = section.Paragraphs[0]

# Reset the image size of the first paragraph
for i in range(paragraph.ChildObjects.Count):
    docObj = paragraph.ChildObjects.get_Item(i)
    if isinstance(docObj, DocPicture):
        picture = DocPicture(docObj)
        picture.Width = 50
        picture.Height = 50
```

---

# Spire.Doc Python Shape Size Reset
## Reset the width and height of a shape in a Word document
```python
doc = Document()

# Get the first section and the first paragraph that contains the shape
section = doc.Sections[0]
para = section.Paragraphs[0]

# Get the second shape and reset the width and height for the shape
shape = para.ChildObjects[1] if isinstance(
    para.ChildObjects[1], ShapeObject) else None
shape.Width = 200
shape.Height = 200
```

---

# spire.doc python shape rotation
## rotate shapes in word document
```python
# Get the first section
section = doc.Sections[0]

# Traverse the word document and set the shape rotation as 20
for i in range(section.Paragraphs.Count):
    para = section.Paragraphs.get_Item(i)
    for j in range(para.ChildObjects.Count):
        obj = para.ChildObjects.get_Item(j)
        if isinstance(obj, ShapeObject):
            (obj if isinstance(obj, ShapeObject) else None).Rotation = 20.0
```

---

# spire.doc python text wrapping
## set text wrapping style for images in Word document
```python
for i in range(doc.Sections.Count):
    sec = doc.Sections.get_Item(i)
    for j in range(sec.Paragraphs.Count):
        para = sec.Paragraphs.get_Item(j)
        pictures = []
        # Get all pictures in the Word document
        for k in range(para.ChildObjects.Count):
            docObj = para.ChildObjects.get_Item(k)
            if docObj.DocumentObjectType == DocumentObjectType.Picture:
                pictures.append(docObj)

        # Set text wrap styles for each picture
        for pic in pictures:
            picture = pic if isinstance(pic, DocPicture) else None
            picture.TextWrappingStyle = TextWrappingStyle.Through
            picture.TextWrappingType = TextWrappingType.Both
```

---

# Spire.Doc Python Image Processing
## Set transparent color for images in document
```python
# Get the first paragraph in the first section
paragraph = doc.Sections[0].Paragraphs[0]

# Set the blue color of the image(s) in the paragraph to transparent
for k in range(paragraph.ChildObjects.Count):
    obj = paragraph.ChildObjects.get_Item(k)
    if isinstance(obj, DocPicture):
        picture = obj if isinstance(obj, DocPicture) else None
        picture.TransparentColor = Color.get_Blue()
```

---

# Update Image in Word Document
## This code demonstrates how to find and replace images in a Word document
```python
# Get all pictures in the Word document
pictures = []
for i in range(doc.Sections.Count):
    sec = doc.Sections.get_Item(i)
    for j in range(sec.Paragraphs.Count):
        para = sec.Paragraphs.get_Item(j)
        for k in range(para.ChildObjects.Count):
            docObj = para.ChildObjects.get_Item(k)
            if docObj.DocumentObjectType == DocumentObjectType.Picture:
                pictures.append(docObj)

# Replace the first picture with a new image file
picture = pictures[0] if isinstance(pictures[0], DocPicture) else None
picture.LoadImage("./Data/E-iceblue.png")
```

---

# Spire.Doc Python Header Management
## Add header only to the first page of a document
```python
# Get the header from the first section
header = doc1.Sections[0].HeadersFooters.Header

# Get the first page header of the destination document
firstPageHeader = doc2.Sections[0].HeadersFooters.FirstPageHeader

# Specify that the current section has a different header/footer for the first page
for i in range(doc2.Sections.Count):
    section = doc2.Sections.get_Item(i)
    section.PageSetup.DifferentFirstPageHeaderFooter = True

# Removes all child objects in firstPageHeader
firstPageHeader.Paragraphs.Clear()

# Add all child objects of the header to firstPageHeader
for j in range(header.ChildObjects.Count):
    obj = header.ChildObjects.get_Item(j)
    firstPageHeader.ChildObjects.Add(obj.Clone())
```

---

# Spire.Doc Python Header Footer Height Adjustment
## Adjust the height of headers and footers in a Word document section
```python
# Get the first section
section = doc.Sections[0]

# Adjust the height of headers in the section
section.PageSetup.HeaderDistance = 100

# Adjust the height of footers in the section
section.PageSetup.FooterDistance = 100
```

---

# Spire.Doc Python Header and Footer Operations
## Copy header from one document to another document
```python
# Get the header section from the source document
header = doc1.Sections[0].HeadersFooters.Header

# Copy each object in the header of source file to destination file
for i in range(doc2.Sections.Count):
    section = doc2.Sections.get_Item(i)
    for j in range(header.ChildObjects.Count):
        obj = header.ChildObjects.get_Item(j)
        section.HeadersFooters.Header.ChildObjects.Add(obj.Clone())
```

---

# spire.doc python header footer
## set different first page header and footer
```python
# Get the section and set the property true
section = doc.Sections[0]
section.PageSetup.DifferentFirstPageHeaderFooter = True

# Set the first page header. Here we append a picture in the header
paragraph1 = section.HeadersFooters.FirstPageHeader.AddParagraph()
paragraph1.Format.HorizontalAlignment = HorizontalAlignment.Right

headerimage = paragraph1.AppendPicture("./Data/E-iceblue.png")

# Set the first page footer
paragraph2 = section.HeadersFooters.FirstPageFooter.AddParagraph()
paragraph2.Format.HorizontalAlignment = HorizontalAlignment.Center
FF = paragraph2.AppendText("First Page Footer")
FF.CharacterFormat.FontSize = 10

# Set the other header & footer. If you only need the first page header & footer, don't set this
paragraph3 = section.HeadersFooters.Header.AddParagraph()
paragraph3.Format.HorizontalAlignment = HorizontalAlignment.Center
NH = paragraph3.AppendText("Spire.Doc for Python")
NH.CharacterFormat.FontSize = 10

paragraph4 = section.HeadersFooters.Footer.AddParagraph()
paragraph4.Format.HorizontalAlignment = HorizontalAlignment.Center
NF = paragraph4.AppendText("E-iceblue")
NF.CharacterFormat.FontSize = 10
```

---

# Spire.Doc Python Header and Footer
## Function to insert header and footer with images and page numbers
```python
def InsertHeaderAndFooter(section):
    header = section.HeadersFooters.Header
    footer = section.HeadersFooters.Footer

    # Insert picture and text to header
    headerParagraph = header.AddParagraph()

    headerPicture = headerParagraph.AppendPicture("./Data/Header.png")

    # Header text
    text = headerParagraph.AppendText("Demo of Spire.Doc")
    text.CharacterFormat.FontName = "Arial"
    text.CharacterFormat.FontSize = 10
    text.CharacterFormat.Italic = True
    headerParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

    # Border
    headerParagraph.Format.Borders.Bottom.BorderType = BorderStyle.Single
    headerParagraph.Format.Borders.Bottom.Space = 0.05

    # Header picture layout - text wrapping
    headerPicture.TextWrappingStyle = TextWrappingStyle.Behind

    # Header picture layout - position
    headerPicture.HorizontalOrigin = HorizontalOrigin.Page
    headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
    headerPicture.VerticalOrigin = VerticalOrigin.Page
    headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top

    # Insert picture to footer
    footerParagraph = footer.AddParagraph()

    footerPicture = footerParagraph.AppendPicture("./Data/Footer.png")

    # Footer picture layout
    footerPicture.TextWrappingStyle = TextWrappingStyle.Behind
    footerPicture.HorizontalOrigin = HorizontalOrigin.Page
    footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
    footerPicture.VerticalOrigin = VerticalOrigin.Page
    footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom

    # Insert page number
    footerParagraph.AppendField("page number", FieldType.FieldPage)
    footerParagraph.AppendText(" of ")
    footerParagraph.AppendField("number of pages", FieldType.FieldNumPages)
    footerParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

    # Border
    footerParagraph.Format.Borders.Top.BorderType = BorderStyle.Single
    footerParagraph.Format.Borders.Top.Space = 0.05
```

---

# spire.doc python header footer
## add image and text to header and footer in word document
```python
# Get the header of the first page
header = doc.Sections[0].HeadersFooters.Header

# Add a paragraph for the header
paragraph = header.AddParagraph()

# Set the format of the paragraph
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Right

# Append a picture in the paragraph
headerimage = paragraph.AppendPicture("./Data/E-iceblue.png")
headerimage.VerticalAlignment = ShapeVerticalAlignment.Bottom

# Get the footer of the first section
footer = doc.Sections[0].HeadersFooters.Footer

# Add a paragraph for the footer
paragraph2 = footer.AddParagraph()

# Set the format of the paragraph
paragraph2.Format.HorizontalAlignment = HorizontalAlignment.Left

# Append a picture in the paragraph
footerimage = paragraph2.AppendPicture("./Data/logo.png")

# Append text in the paragraph
TR = paragraph2.AppendText(
    "Copyright © 2013 e-iceblue. All Rights Reserved.")
TR.CharacterFormat.FontName = "Arial"
TR.CharacterFormat.FontSize = 10
TR.CharacterFormat.TextColor = Color.get_Black()
```

---

# Spire.Doc Header Protection
## Lock header in Word document by protecting document while allowing form fields
```python
# Get the first section
section = doc.Sections[0]

# Protect the document and set the ProtectionType as AllowOnlyFormFields
doc.Protect(ProtectionType.AllowOnlyFormFields, "123")

# Set the ProtectForm as false to unprotect the section
section.ProtectForm = False
```

---

# spire.doc python headers and footers
## create different headers and footers for odd and even pages
```python
# Get the section
section = doc.Sections[0]

# Set the DifferentOddAndEvenPagesHeaderFooter property to true
section.PageSetup.DifferentOddAndEvenPagesHeaderFooter = True

# Add odd header
oddHeaderParagraph = section.HeadersFooters.OddHeader.AddParagraph()
oddHeaderText = oddHeaderParagraph.AppendText("Odd Header")
oddHeaderParagraph.Format.HorizontalAlignment = HorizontalAlignment.Center
oddHeaderText.CharacterFormat.FontName = "Arial"
oddHeaderText.CharacterFormat.FontSize = 10

# Add even header
evenHeaderParagraph = section.HeadersFooters.EvenHeader.AddParagraph()
evenHeaderText = evenHeaderParagraph.AppendText("Even Header from E-iceblue Using Spire.Doc")
evenHeaderParagraph.Format.HorizontalAlignment = HorizontalAlignment.Center
evenHeaderText.CharacterFormat.FontName = "Arial"
evenHeaderText.CharacterFormat.FontSize = 10

# Add odd footer
oddFooterParagraph = section.HeadersFooters.OddFooter.AddParagraph()
oddFooterText = oddFooterParagraph.AppendText("Odd Footer")
oddFooterParagraph.Format.HorizontalAlignment = HorizontalAlignment.Center
oddFooterText.CharacterFormat.FontName = "Arial"
oddFooterText.CharacterFormat.FontSize = 10

# Add even footer
evenFooterParagraph = section.HeadersFooters.EvenFooter.AddParagraph()
evenFooterText = evenFooterParagraph.AppendText("Even Footer from E-iceblue Using Spire.Doc")
evenFooterText.CharacterFormat.FontName = "Arial"
evenFooterText.CharacterFormat.FontSize = 10
evenFooterParagraph.Format.HorizontalAlignment = HorizontalAlignment.Center
```

---

# Spire.Doc Page Border Surround Configuration
## Configure page borders and their interaction with headers and footers
```python
# Create a new document
doc = Document()
section = doc.AddSection()

# Add a sample page border to the document
section.PageSetup.Borders.BorderType = BorderStyle.Wave
section.PageSetup.Borders.Color = Color.get_Green()
section.PageSetup.Borders.Left.Space = 20.0
section.PageSetup.Borders.Right.Space = 20.0

# Add a header and set its format
paragraph1 = section.HeadersFooters.Header.AddParagraph()
paragraph1.Format.HorizontalAlignment = HorizontalAlignment.Right
headerText = paragraph1.AppendText("Header isn't included in page border")
headerText.CharacterFormat.FontName = "Calibri"
headerText.CharacterFormat.FontSize = 20.0
headerText.CharacterFormat.Bold = True

# Add a footer and set its format
paragraph2 = section.HeadersFooters.Footer.AddParagraph()
paragraph2.Format.HorizontalAlignment = HorizontalAlignment.Left
footerText = paragraph2.AppendText("Footer is included in page border")
footerText.CharacterFormat.FontName = "Calibri"
footerText.CharacterFormat.FontSize = 20.0
footerText.CharacterFormat.Bold = True

# Set the header not included in the page border while the footer included
section.PageSetup.PageBorderIncludeHeader = False
section.PageSetup.HeaderDistance = 40.0
section.PageSetup.PageBorderIncludeFooter = True
section.PageSetup.FooterDistance = 40.0
```

---

# Remove Footer from Word Document
## This code removes all types of footers (first page, odd page, even page) from a Word document section
```python
# Get the first section
section = doc.Sections[0]

# Clear footer in the first page
footer = section.HeadersFooters[HeaderFooterType.FooterFirstPage]
if footer is not None:
    footer.ChildObjects.Clear()

# Clear footer in the odd page
footer = section.HeadersFooters[HeaderFooterType.FooterOdd]
if footer is not None:
    footer.ChildObjects.Clear()

# Clear footer in the even page
footer = section.HeadersFooters[HeaderFooterType.FooterEven]
if footer is not None:
    footer.ChildObjects.Clear()
```

---

# spire.doc python header removal
## Remove headers from Word document
```python
# Get the first section of the document
section = doc.Sections[0]

# Traverse the word document and clear all headers in different type
for i in range(section.Paragraphs.Count):
    para = section.Paragraphs.get_Item(i)
    for j in range(para.ChildObjects.Count):
        obj = para.ChildObjects.get_Item(j)
        # Clear header in the first page
        header = None
        header = section.HeadersFooters[HeaderFooterType.HeaderFirstPage]
        if header is not None:
            header.ChildObjects.Clear()
        # Clear header in the odd page
        header = section.HeadersFooters[HeaderFooterType.HeaderOdd]
        if header is not None:
            header.ChildObjects.Clear()
        # Clear header in the even page
        header = section.HeadersFooters[HeaderFooterType.HeaderEven]
        if header is not None:
            header.ChildObjects.Clear()
```

---

# Add Alternative Text for Table
## This code demonstrates how to add alternative text (title and description) to a table in a Word document
```python
# Get the first section
section = doc.Sections[0]

# Get the first table in the section
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Add alternative text
# Add title
table.Title = "Table 1"
# Add description
table.TableDescription = "Description Text"
```

---

# spire.doc python table
## Add or delete rows in a Word table
```python
section = document.Sections[0]
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Delete the seventh row
table.Rows.RemoveAt(7)

# Add a row and insert it into specific position
row = TableRow(document)
for i in range(table.Rows[0].Cells.Count):
    tc = row.AddCell()
    paragraph = tc.AddParagraph()
    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
    paragraph.AppendText("Added")
table.Rows.Insert(2, row)
# Add a row at the end of table
table.AddRow()
```

---

# Spire.Doc Python Table Operations
## Add or remove columns from a table in a Word document
```python
def AddColumn(table, columnIndex):
    for r in range(table.Rows.Count):
        addCell = TableCell(table.Document)
        table.Rows[r].Cells.Insert(columnIndex, addCell)

def RemoveColumn(table, columnIndex):
    for r in range(table.Rows.Count):
        table.Rows[r].Cells.RemoveAt(columnIndex)
```

---

# Spire.Doc Python Table Operations
## Add picture to table cell
```python
# Get the first table from the first section of the document
table1 = doc.Sections[0].Tables[0]

# Add a picture to the specified table cell and set picture size
picture = table1.Rows[1].Cells[2].Paragraphs[0].AppendPicture("./Data/Spire.Doc.png")

picture.Width = 100
picture.Height = 100
```

---

# Spire.Doc Python Table Creation
## Add table to Word document using data table
```python
@staticmethod
def _FillTableUsingDataTable(table, dataTable):
    columnCount = len(dataTable[0])

    for dataRow in dataTable:
        row = table.AddRow(columnCount)
        i = 0
        for col in dataRow:
            #columnIndex = dataTable.Columns.IndexOf(dataColumn)
            value = str(col.text)
            cell = row.Cells.get_Item(i)
            paragraph = cell.AddParagraph()
            paragraph.AppendText(value)
            #Set the alignment of cell
            cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
            paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
            i += 1

#Create a Word document
document = Document()

#Get the first section
section = document.AddSection()

#Add paragraph style
style = ParagraphStyle(document)
style.CharacterFormat.FontSize = 20
style.CharacterFormat.Bold = True
style.CharacterFormat.TextColor = Color.get_CadetBlue()
document.Styles.Add(style)

#Create a paragraph and append text
para = section.AddParagraph()
para.AppendText("Table")
#Apply style
para.Format.HorizontalAlignment = HorizontalAlignment.Center
para.ApplyStyle(style.Name)

#Add a table
table = section.AddTable(True)
#Set its width
table.PreferredWidth = PreferredWidth(WidthType.Percentage, 100)

#Fill table with the data of datatable
# _FillTableUsingDataTable(table, dataTable)

#Set table style
table.Format.Paddings.SetAll(5)
row = table.FirstRow
i = 0
while i < row.Cells.Count:
    row.Cells.get_Item(i).CellFormat.Shading.BackgroundPatternColor = Color.get_CadetBlue()
    i += 1
```

---

# Spire.Doc Python Table Formatting
## Allow table rows to break across pages
```python
section = document.Sections[0]
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

for i in range(table.Rows.Count):
    row = table.Rows.get_Item(i)
    # Allow break across pages
    row.RowFormat.IsBreakAcrossPages = True
```

---

# Spire.Doc Python Table AutoFit
## Auto-fit table to contents in Word document
```python
section = document.Sections[0]
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Automatically fit the table to the cell content
table.AutoFit(AutoFitBehaviorType.AutoFitToContents)
```

---

# spire.doc python table auto-fit
## set table to fixed column widths
```python
# Get the first table in the first section
section = document.Sections[0]
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None
# The table is set to a fixed size
table.AutoFit(AutoFitBehaviorType.FixedColumnWidths)
```

---

# Spire.Doc Python Table AutoFit
## Automatically fit table to window width
```python
# Get table from document
section = document.Sections[0]
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Automatically fit the table to the active window width
table.AutoFit(AutoFitBehaviorType.AutoFitToWindow)
```

---

# spire.doc python table cell merge status
## Check the merge status of cells in a table
```python
# Get the first section
section = doc.Sections[0]

# Get the first table in the section
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

content = ''
for i in range(table.Rows.Count):
    tableRow = table.Rows[i]
    for j in range(tableRow.Cells.Count):
        tableCell = tableRow.Cells[j]
        verticalMerge = tableCell.CellFormat.VerticalMerge
        horizontalMerge = tableCell.GridSpan
        if verticalMerge == CellMerge.none and horizontalMerge == 1:
            content += "Row " + str(i) + ", cell " + str(j) + ": "
            content += "This cell isn't merged."
            content += "\n"
        else:
            content += "Row " + str(i) + ", cell " + str(j) + ": "
            content += "This cell is merged."
            content += "\n"
    content += "\n"
```

---

# Spire.Doc Python Table Row Cloning
## Clone a table row and add it to the table
```python
# Get the first section
se = doc.Sections[0]

# Get the first row of the first table
firstRow = se.Tables[0].Rows[0]

# Copy the first row to clone_FirstRow via TableRow.clone()
clone_FirstRow = firstRow.Clone()

se.Tables[0].Rows.Add(clone_FirstRow)
```

---

# Spire.Doc Python Table Operations
## Clone a table in a Word document
```python
# Get the first section
se = doc.Sections[0]

# Get the first table
original_Table = se.Tables[0]

# Copy the existing table to copied_Table via Table.clone()
copied_Table = original_Table.Clone()

# Get the last row of table
lastRow = copied_Table.Rows[copied_Table.Rows.Count - 1]

# Change last row data
i = 0
while i < lastRow.Cells.Count - 1:
    lastRow.Cells[i].Paragraphs[0].Text = "New text"
    i += 1
    
# Add copied_Table in section
se.Tables.Add(copied_Table)
```

---

# Spire.Doc Python Table Operations
## Combine and split tables in Word documents
```python
# Split a table
def SplitTable():
    #Get the first section
    section = doc.Sections[0]

    #Get the first table
    table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

    #We will split the table at the third row
    splitIndex = 2

    #Create a new table for the split table
    newTable = Table(section.Document)

    #Add rows to the new table
    for i in range(splitIndex, table.Rows.Count):
        newTable.Rows.Add(table.Rows[i].Clone())

    #Remove rows from original table
    for i in range(table.Rows.Count - 1, splitIndex - 1, -1):
        table.Rows.RemoveAt(i)

    #Add the new table in section
    section.Tables.Add(newTable)
```

## Combine tables in Word document
```python
def CombineTables():
    #Get the first section
    section = doc.Sections[0]

    #Get the first and second table
    table1 = section.Tables[0] if isinstance(section.Tables[0], Table) else None
    table2 = section.Tables[1] if isinstance(section.Tables[1], Table) else None

    #Add the rows of table2 to table1
    for i in range(table2.Rows.Count):
        table1.Rows.Add(table2.Rows[i].Clone())

    #Remove the table2
    section.Tables.Remove(table2)
```

---

# Spire.Doc for Python - Create Nested Table
## Demonstrates how to create a nested table within a table cell in a Word document

```python
# Create a new document
doc = Document()
section = doc.AddSection()

# Add a table
table = section.AddTable(True)
table.ResetCells(2, 2)

# Set column width
table.Rows[0].Cells[0].SetCellWidth(70, CellWidthType.Point)
table.Rows[1].Cells[0].SetCellWidth(70, CellWidthType.Point)
table.AutoFit(AutoFitBehaviorType.AutoFitToWindow)

# Add a nested table to cell(first row, second column)
nestedTable = table.Rows[0].Cells[1].AddTable(True)
nestedTable.ResetCells(4, 3)
nestedTable.AutoFit(AutoFitBehaviorType.AutoFitToContents)
```

---

# Spire.Doc Python table creation
## Create and format a table in a Word document
```python
def addTable(section):
    table = section.AddTable(True)
    table.ResetCells(len(data) + 1, len(header))

    # ***************** First Row *************************
    row = table.Rows[0]
    row.IsHeader = True
    row.Height = 20 #unit: point, 1point = 0.3528 mm
    row.HeightType = TableRowHeightType.Exactly
    i = 0
    while i < row.Cells.Count:
        row.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.get_Gray()
        i += 1
    i = 0
    while i < len(header):
        row.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle
        p = row.Cells[i].AddParagraph()
        p.Format.HorizontalAlignment = HorizontalAlignment.Center
        txtRange = p.AppendText(header[i])
        txtRange.CharacterFormat.Bold = True
        i += 1

    r = 0
    while r < len(data):
        dataRow = table.Rows[r + 1]
        dataRow.Height = 20
        dataRow.HeightType = TableRowHeightType.Exactly
        i = 0
        while i < dataRow.Cells.Count:
            dataRow.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.Empty()
            i += 1
        c = 0
        while c < len(data[r]):
            dataRow.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle
            dataRow.Cells[c].AddParagraph().AppendText(data[r][c])
            c += 1
        r += 1

    for j in range(1, table.Rows.Count):
        if math.fmod(j, 2) == 0:
            row2 = table.Rows[j]
            for f in range(row2.Cells.Count):
                row2.Cells[f].CellFormat.Shading.BackgroundPatternColor = Color.get_LightBlue()
```

---

# Spire.Doc Python Table Creation
## Creating a table directly in a Word document with formatting
```python
# Create a Word document
doc = Document()

# Add a section
section = doc.AddSection()

#Create a table 
table = Table(doc)
table.ResetCells(1,2)

#Set the width of table
table.PreferredWidth = PreferredWidth(WidthType.Percentage, int(100))

#Set the border of table
table.Format.Borders.BorderType = BorderStyle.Single

#Create a table row
row = table.Rows[0]
row.Height = 50.0

#Create a table cell
cell1 = table.Rows[0].Cells[0]

#Add a paragraph
para1 = cell1.AddParagraph()
#Append text in the paragraph
para1.AppendText("Row 1, Cell 1")
#Set the horizontal alignment of paragrah
para1.Format.HorizontalAlignment = HorizontalAlignment.Center

#Set the background color of cell
cell1.CellFormat.Shading.BackgroundPatternColor = Color.get_CadetBlue()

#Set the vertical alignment of paragraph
cell1.CellFormat.VerticalAlignment = VerticalAlignment.Middle

#Create a table cell
cell2 = table.Rows[0].Cells[1]
para2 = cell2.AddParagraph()
para2.AppendText("Row 1, Cell 2")
para2.Format.HorizontalAlignment = HorizontalAlignment.Center
cell2.CellFormat.Shading.BackgroundPatternColor = Color.get_CadetBlue()
cell2.CellFormat.VerticalAlignment = VerticalAlignment.Middle
row.Cells.Add(cell2)

#Add the table in the section
section.Tables.Add(table)
```

---

# Spire.Doc Python HTML to Table
## Create a table in Word document from HTML
```python
# Create a Word document
document = Document()

# Add a section
section = document.AddSection()

# Add a paragraph and append html string
html = "<table><tr><td>Cell 1</td><td>Cell 2</td></tr></table>"
section.AddParagraph().AppendHTML(html)
```

---

# spire.doc python vertical table
## create a vertical table in Word document
```python
# Create Word document.
document = Document()

# Add a new section.
section = document.AddSection()

# Add a table with rows and columns and set the text for the table.
table = section.AddTable()
table.ResetCells(1, 1)
cell = table.Rows[0].Cells[0]
table.Rows[0].Height = 150
cell.AddParagraph().AppendText("Draft copy in vertical style")

# Set the TextDirection for the table to RightToLeftRotated.
cell.CellFormat.TextDirection = TextDirection.RightToLeftRotated

# Set the table format.
table.Format.WrapTextAround = True
table.Format.Positioning.VertRelationTo = VerticalRelation.Page
table.Format.Positioning.HorizRelationTo = HorizontalRelation.Page
table.Format.Positioning.HorizPosition = section.PageSetup.PageSize.Width - table.Width
table.Format.Positioning.VertPosition = 200
```

---

# spire.doc python table borders
## Set different borders for table and cells
```python
def setTableBorders(table):
    table.Format.Borders.BorderType=BorderStyle.Single
    table.Format.Borders.LineWidth=3.0
    table.Format.Borders.Color=Color.get_Red()

def setCellBorders(tableCell):
    tableCell.CellFormat.Borders.BorderType=BorderStyle.DotDash
    tableCell.CellFormat.Borders.LineWidth=1.0
    tableCell.CellFormat.Borders.Color=Color.get_Green()
```

---

# Spire.Doc Python Table Formatting
## Format merged cells in a Word document table
```python
# Create word document
document = Document()

# Create a new section
section = document.AddSection()

# Create a table
table = section.AddTable(True)
table.ResetCells(4, 3)

# Create a new style for merged cells
style = ParagraphStyle(document)
style.Name = "Style"
style.CharacterFormat.TextColor = Color.get_DeepSkyBlue()
style.CharacterFormat.Italic = True
style.CharacterFormat.Bold = True
style.CharacterFormat.FontSize = 13
document.Styles.Add(style)

# Merge cell horizontally
table.ApplyHorizontalMerge(0, 0, 1)
# Apply style to horizontally merged cell
table.Rows[0].Cells[0].Paragraphs[0].ApplyStyle(style.Name)
# Set vertical and horizontal alignment for horizontally merged cell
table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle
table.Rows[0].Cells[0].Paragraphs[0].Format.HorizontalAlignment = HorizontalAlignment.Center

# Merge cell vertically
table.ApplyVerticalMerge(0, 1, 3)
# Apply style to vertically merged cell
table.Rows[1].Cells[0].Paragraphs[0].ApplyStyle(style.Name)
# Set vertical and horizontal alignment for vertically merged cell
table.Rows[1].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle
table.Rows[1].Cells[0].Paragraphs[0].Format.HorizontalAlignment = HorizontalAlignment.Left
# Set column width for vertically merged cell
table.Rows[1].Cells[0].SetCellWidth(20, CellWidthType.Percentage)
```

---

# Spire.Doc Python Table Diagonal Border
## Get diagonal border properties from a table cell in a Word document
```python
# Get the first section
section = doc.Sections[0]

# Get the first table in the section
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Get the setting of the diagonal border of table cell
bs_UP = table.Rows[0].Cells[0].CellFormat.Borders.DiagonalUp.BorderType
color_UP = table.Rows[0].Cells[0].CellFormat.Borders.DiagonalUp.Color
width_UP = table.Rows[0].Cells[0].CellFormat.Borders.DiagonalUp.LineWidth
bs_Down = table.Rows[0].Cells[0].CellFormat.Borders.DiagonalDown.BorderType
color_Down = table.Rows[0].Cells[0].CellFormat.Borders.DiagonalDown.Color
width_Down = table.Rows[0].Cells[0].CellFormat.Borders.DiagonalDown.LineWidth
```

---

# spire.doc python table index
## get table, row and cell indices from word document
```python
# Get the first section
section = doc.Sections[0]

# Get the first table in the section
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Get table collections
collections = section.Tables

# Get the table index
tableIndex = collections.IndexOf(table)

# Get the index of the last table row
row = table.LastRow
rowIndex = row.GetRowIndex()

# Get the index of the last table cell
cell = row.LastChild if isinstance(row.LastChild, TableCell) else None
cellIndex = cell.GetCellIndex()
```

---

# spire.doc python table position
## get table position information from a Word document
```python
# Get the first section
section = document.Sections[0]
# Get the first table
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Verify whether the table uses "Around" text wrapping or not.
if table.Format.WrapTextAround:
    positon = table.Format.Positioning

    # Horizontal position
    horizPosition = positon.HorizPosition
    horizPositionAbs = positon.HorizPositionAbs
    horizRelationTo = positon.HorizRelationTo
    
    # Vertical position
    vertPosition = positon.VertPosition
    vertPositionAbs = positon.VertPositionAbs
    vertRelationTo = positon.VertRelationTo
    
    # Distance from surrounding text
    distanceFromTop = positon.DistanceFromTop
    distanceFromLeft = positon.DistanceFromLeft
    distanceFromBottom = positon.DistanceFromBottom
    distanceFromRight = positon.DistanceFromRight
```

---

# Spire.Doc Python Table Cell Operations
## Merge and split table cells in a Word document
```python
# The method shows how to merge cell horizontally
table.ApplyHorizontalMerge(6, 2, 3)
# The method shows how to merge cell vertically
table.ApplyVerticalMerge(2, 4, 5)
# The method shows how to split the cell
table.Rows[8].Cells[3].SplitCell(2, 2)
```

---

# Spire.Doc Python Table Formatting
## Modify table, row and cell formats in Word documents

```python
# Modify table format
def _MoidyTableFormat(table):
    # Set table width
    table.PreferredWidth = PreferredWidth(WidthType.Twip, int(6000))

    # Apply style for table
    table.ApplyStyle(DefaultTableStyle.ColorfulGridAccent3)

    # Set table padding
    table.Format.Paddings.SetAll(5)

    # Set table title and description
    table.Title = "Spire.Doc for Python"
    table.TableDescription = "Spire.Doc for Python is a professional Word Python library"


# Modify row format
def _ModifyRowFormat(table):
    # Set cell spacing
    table.Format.CellSpacing = 2
    # Set row height
    table.Rows[1].HeightType = TableRowHeightType.Exactly
    table.Rows[1].Height = 20
    row2 = table.Rows[2]
    # Set background color
    i = 0
    while i < row2.Cells.Count:
        row2.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.get_DarkSeaGreen()
        i += 1


# Modify cell format
def _ModifyCellFormat(table):
    # Set alignment
    table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle
    table.Rows[0].Cells[0].Paragraphs[0].Format.HorizontalAlignment = HorizontalAlignment.Center
    # Set background color
    table.Rows[1].Cells[0].CellFormat.Shading.BackgroundPatternColor = Color.get_DarkSeaGreen()
    # Set cell border
    table.Rows[2].Cells[0].CellFormat.Borders.BorderType=BorderStyle.Single
    table.Rows[2].Cells[0].CellFormat.Borders.LineWidth=1
    table.Rows[2].Cells[0].CellFormat.Borders.Left.Color = Color.get_Red()
    table.Rows[2].Cells[0].CellFormat.Borders.Right.Color = Color.get_Red()
    table.Rows[2].Cells[0].CellFormat.Borders.Top.Color = Color.get_Red()
    table.Rows[2].Cells[0].CellFormat.Borders.Bottom.Color = Color.get_Red()
    # Set text direction
    table.Rows[3].Cells[0].CellFormat.TextDirection = TextDirection.RightToLeft
```

---

# Spire.Doc Python Table
## Prevent page breaks in a table
```python
# Get the table from Word document.
table = document.Sections[0].Tables[0] if isinstance(
    document.Sections[0].Tables[0], Table) else None

# Change the paragraph setting to keep them together.
for i in range(table.Rows.Count):
    row = table.Rows.get_Item(i)
    for j in range(row.Cells.Count):
        cell = row.Cells.get_Item(j)
        for k in range(cell.Paragraphs.Count):
            p = cell.Paragraphs.get_Item(k)
            p.Format.KeepFollow = True
```

---

# Spire.Doc Python Table Operations
## Remove a table from a document
```python
# Remove the first Table from the first section of the document
doc.Sections[0].Tables.RemoveAt(0)
```

---

# Spire.Doc Python Table Header
## Create table with header rows that repeat on each page
```python
# Create word document
document = Document()

# Create a new section
section = document.AddSection()

# Create a table with default borders
table = section.AddTable(True)
# Set table width to 100%
width = PreferredWidth(WidthType.Percentage, 100)
table.PreferredWidth = width

# Add a new row
row = table.AddRow()
# Set the row as a table header (this makes it repeat on each page)
row.IsHeader = True

# Add a new cell for row
cell = row.AddCell()

# Set the backcolor of row
i = 0
while i < row.Cells.Count:
    row.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.get_LightGray()
    i += 1

cell.SetCellWidth(100, CellWidthType.Percentage)
# Add a paragraph for cell to put some data
parapraph = cell.AddParagraph()
# Add text
parapraph.AppendText("Row Header 1")
# Set paragraph horizontal center alignment
parapraph.Format.HorizontalAlignment = HorizontalAlignment.Center

row = table.AddRow(False, 1)
row.IsHeader = True

i = 0
while i < row.Cells.Count:
    row.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.get_Ivory()
    i += 1

# Set row height
row.Height = 30
cell = row.Cells[0]
cell.SetCellWidth(100, CellWidthType.Percentage)
# Set cell vertical middle alignment
cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
# Add a paragraph for cell to put some data
parapraph = cell.AddParagraph()
# Add text
parapraph.AppendText("Row Header 2")
parapraph.Format.HorizontalAlignment = HorizontalAlignment.Center
```

---

# Spire.Doc Python Table Text Replacement
## Replace text in table using regex and direct string replacement
```python
# Get the first section
section = doc.Sections[0]

# Get the first table in the section
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Define a regular expression to match the {} with its content
regex = Regex("""{[^\\}]+\\}""")

# Replace the text of table with regex
table.Replace(regex, "E-iceblue")

# Replace old text with new text in table
table.Replace("Beijing", "Component", False, True)
```

---

# spire.doc python table column width
## set column width in a table
```python
# Get section from document
section = document.Sections[0]
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

# Traverse the first column
for i in range(table.Rows.Count):
    # Set the width and type of the cell
    table.Rows[i].Cells[0].SetCellWidth(200, CellWidthType.Point)
```

---

# Spire.Doc Python Table Positioning
## Set table outside position relative to image in document header
```python
# Create a new word document and add new section
doc = Document()
sec = doc.AddSection()

# Get header
header = doc.Sections[0].HeadersFooters.Header

# Add new paragraph on header and set HorizontalAlignment of the paragraph as left
paragraph = header.AddParagraph()
paragraph.Format.HorizontalAlignment = HorizontalAlignment.Left

# Load an image for the paragraph
headerimage = paragraph.AppendPicture(inputFile)

# Add a table of 4 rows and 2 columns
table = header.AddTable()
table.ResetCells(4, 2)

# Set the position of the table to the right of the image
table.Format.WrapTextAround = True
table.Format.Positioning.HorizPositionAbs = HorizontalPosition.Outside
table.Format.Positioning.VertRelationTo = VerticalRelation.Margin
table.Format.Positioning.VertPosition = 43

# Add contents for the table
for r in range(0, 4):
    dataRow = table.Rows[r]
    for c in range(0, 2):
        if c == 0:
            par = dataRow.Cells[c].AddParagraph()
            par.AppendText("Spire.Doc.left")
            par.Format.HorizontalAlignment = HorizontalAlignment.Left
            dataRow.Cells[c].SetCellWidth(180,CellWidthType.Point)
        else:
            par = dataRow.Cells[c].AddParagraph()
            par.AppendText("Spire XLS.right")
            par.Format.HorizontalAlignment = HorizontalAlignment.Right
            dataRow.Cells[c].SetCellWidth(180,CellWidthType.Point)
```

---

# Spire.Doc Python Table
## Set table style and borders
```python
# Apply the table style
table.ApplyStyle(DefaultTableStyle.ColorfulList)

# Set right border of table
table.Format.Borders.Right.BorderType = BorderStyle.Hairline
table.Format.Borders.Right.LineWidth = 1.0
table.Format.Borders.Right.Color = Color.get_Red()

# Set top border of table
table.Format.Borders.Top.BorderType = BorderStyle.Hairline
table.Format.Borders.Top.LineWidth = 1.0
table.Format.Borders.Top.Color = Color.get_Green()

# Set left border of table
table.Format.Borders.Left.BorderType = BorderStyle.Hairline
table.Format.Borders.Left.LineWidth = 1.0
table.Format.Borders.Left.Color = Color.get_Yellow()

# Set bottom border is none
table.Format.Borders.Bottom.BorderType = BorderStyle.DotDash

# Set vertical and horizontal border
table.Format.Borders.Vertical.BorderType = BorderStyle.Dot
table.Format.Borders.Horizontal.BorderType = BorderStyle.none
table.Format.Borders.Vertical.Color = Color.get_Orange()
```

---

# Spire.Doc Python Table Vertical Alignment
## Set vertical alignment for table cells in a Word document
```python
# Create a new Word document and add a new section
doc = Document()
section = doc.AddSection()

# Add a table with 3 columns and 3 rows
table = section.AddTable(True)
table.ResetCells(3, 3)

# Merge rows
table.ApplyVerticalMerge(0, 0, 2)

# Set the vertical alignment for each cell, default is top
table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle
table.Rows[0].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Top
table.Rows[0].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Top
table.Rows[1].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle
table.Rows[1].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle
table.Rows[2].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Bottom
table.Rows[2].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Bottom
```

---

# spire.doc python image hyperlink
## create an image hyperlink in a document
```python
# Create a document
doc = Document()

# Add a section and paragraph
section = doc.Sections[0]
paragraph = section.AddParagraph()

# Create a DocPicture object and load an image
picture = DocPicture(doc)
picture.LoadImage("./Data/Spire.Doc.png")

# Add an image hyperlink to the paragraph
paragraph.AppendHyperlink(
    "https://www.e-iceblue.com/Introduce/doc-for-python.html", picture, HyperlinkType.WebLink)
```

---

# spire.doc python hyperlink finder
## find and extract hyperlinks from a document
```python
# Create a hyperlink list
hyperlinks = []
hyperlinksText = ''
# Iterate through the items in the sections to find all hyperlinks
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    for j in range(section.Body.ChildObjects.Count):
        sec = section.Body.ChildObjects.get_Item(j)
        if sec.DocumentObjectType == DocumentObjectType.Paragraph:
            for k in range((sec if isinstance(sec, Paragraph) else None).ChildObjects.Count):
                para = (sec if isinstance(sec, Paragraph)
                        else None).ChildObjects.get_Item(k)
                if para.DocumentObjectType == DocumentObjectType.Field:
                    field = para if isinstance(para, Field) else None
                    if field.Type == FieldType.FieldHyperlink:
                        hyperlinks.append(field)
                        # Get the hyperlink text
                        hyperlinksText += field.FieldText + "\r\n"
```

---

# Spire.Doc Python Hyperlink Creation
## Create various types of hyperlinks in a Word document
```python
def _InsertHyperlink(section):
    paragraph = section.Paragraphs[0] if section.Paragraphs.Count > 0 else section.AddParagraph()
    paragraph.AppendText(
        "Spire.Doc for Python \r\n e-iceblue company Ltd. 2002-2010 All rights reserverd")
    paragraph.ApplyStyle(BuiltinStyle.Heading2)

    paragraph = section.AddParagraph()
    paragraph.AppendText("Home page")
    paragraph.ApplyStyle(BuiltinStyle.Heading2)
    paragraph = section.AddParagraph()
    paragraph.AppendHyperlink(
        "www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink)

    paragraph = section.AddParagraph()
    paragraph.AppendText("Contact US")
    paragraph.ApplyStyle(BuiltinStyle.Heading2)
    paragraph = section.AddParagraph()
    paragraph.AppendHyperlink(
        "mailto:support@e-iceblue.com", "support@e-iceblue.com", HyperlinkType.EMailLink)

    paragraph = section.AddParagraph()
    paragraph.AppendText("Forum")
    paragraph.ApplyStyle(BuiltinStyle.Heading2)
    paragraph = section.AddParagraph()
    paragraph.AppendHyperlink(
        "www.e-iceblue.com/forum/", "www.e-iceblue.com/forum/", HyperlinkType.WebLink)

    paragraph = section.AddParagraph()
    paragraph.AppendText("Download Link")
    paragraph.ApplyStyle(BuiltinStyle.Heading2)
    paragraph = section.AddParagraph()
    paragraph.AppendHyperlink("www.e-iceblue.com/Download/download-word-for-python-now.html",
                              "www.e-iceblue.com/Download/download-word-for-python-now.html", HyperlinkType.WebLink)

    paragraph = section.AddParagraph()
    paragraph.AppendText("Insert Link On Image")
    paragraph.ApplyStyle(BuiltinStyle.Heading2)
    paragraph = section.AddParagraph()
    picture = paragraph.AppendPicture("./Data/Spire.Doc.png")

    paragraph.AppendHyperlink(
        "www.e-iceblue.com/Introduce/doc-for-python.html", picture, HyperlinkType.WebLink)
```

---

# Spire.Doc Python Modify Hyperlink
## Find and modify hyperlink text in a Word document
```python
# Find all hyperlinks in the Word document
hyperlinks = []
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    for j in range(section.Body.ChildObjects.Count):
        sec = section.Body.ChildObjects.get_Item(j)
        if sec.DocumentObjectType == DocumentObjectType.Paragraph:
            for k in range((sec if isinstance(sec, Paragraph) else None).ChildObjects.Count):
                para = (sec if isinstance(sec, Paragraph)
                        else None).ChildObjects.get_Item(k)
                if para.DocumentObjectType == DocumentObjectType.Field:
                    field = para if isinstance(para, Field) else None

                    if field.Type == FieldType.FieldHyperlink:
                        hyperlinks.append(field)

# Reset the property of hyperlinks[0].FieldText by using the index of the hyperlink
hyperlinks[0].FieldText = "Spire.Doc component"
```

---

# spire.doc python hyperlink removal
## remove hyperlinks from word document while preserving text
```python
def _FindAllHyperlinks(document):
    hyperlinks = []
    # Iterate through the items in the sections to find all hyperlinks
    for i in range(document.Sections.Count):
        section = document.Sections.get_Item(i)
        for j in range(section.Body.ChildObjects.Count):
            sec = section.Body.ChildObjects.get_Item(j)
            if sec.DocumentObjectType == DocumentObjectType.Paragraph:
                for k in range((sec if isinstance(sec, Paragraph) else None).ChildObjects.Count):
                    para = (sec if isinstance(sec, Paragraph)
                            else None).ChildObjects.get_Item(k)
                    if para.DocumentObjectType == DocumentObjectType.Field:
                        field = para if isinstance(para, Field) else None
                        if field.Type == FieldType.FieldHyperlink:
                            hyperlinks.append(field)
    return hyperlinks

# Flatten the hyperlink field
def _FlattenHyperlinks(field):
    ownerParaIndex = field.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(
        field.OwnerParagraph)
    fieldIndex = field.OwnerParagraph.ChildObjects.IndexOf(field)
    sepOwnerPara = field.Separator.OwnerParagraph
    sepOwnerParaIndex = field.Separator.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(
        field.Separator.OwnerParagraph)
    sepIndex = field.Separator.OwnerParagraph.ChildObjects.IndexOf(
        field.Separator)
    endIndex = field.End.OwnerParagraph.ChildObjects.IndexOf(field.End)
    endOwnerParaIndex = field.End.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(
        field.End.OwnerParagraph)

    _FormatFieldResultText(field.Separator.OwnerParagraph.OwnerTextBody,
                           sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex)

    field.End.OwnerParagraph.ChildObjects.RemoveAt(endIndex)

    for i in range(sepOwnerParaIndex, ownerParaIndex - 1, -1):
        if i == sepOwnerParaIndex and i == ownerParaIndex:
            for j in range(sepIndex, fieldIndex - 1, -1):
                field.OwnerParagraph.ChildObjects.RemoveAt(j)

        elif i == ownerParaIndex:
            for j in range(field.OwnerParagraph.ChildObjects.Count - 1, fieldIndex - 1, -1):
                field.OwnerParagraph.ChildObjects.RemoveAt(j)

        elif i == sepOwnerParaIndex:
            for j in range(sepIndex, -1, -1):
                sepOwnerPara.ChildObjects.RemoveAt(j)
        else:
            field.OwnerParagraph.OwnerTextBody.ChildObjects.RemoveAt(i)

# Remove the font color and underline format of the hyperlinks
def _FormatFieldResultText(ownerBody, sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex):
    for i in range(sepOwnerParaIndex, endOwnerParaIndex + 1):
        para = ownerBody.ChildObjects[i] if isinstance(
            ownerBody.ChildObjects[i], Paragraph) else None
        if i == sepOwnerParaIndex and i == endOwnerParaIndex:
            for j in range(sepIndex + 1, endIndex):
                _FormatText(para.ChildObjects[j] if isinstance(
                    para.ChildObjects[j], TextRange) else None)

        elif i == sepOwnerParaIndex:
            for j in range(sepIndex + 1, para.ChildObjects.Count):
                _FormatText(para.ChildObjects[j] if isinstance(
                    para.ChildObjects[j], TextRange) else None)
        elif i == endOwnerParaIndex:
            for j in range(0, endIndex):
                _FormatText(para.ChildObjects[j] if isinstance(
                    para.ChildObjects[j], TextRange) else None)
        else:
            for j, unusedItem in enumerate(para.ChildObjects):
                _FormatText(para.ChildObjects[j] if isinstance(
                    para.ChildObjects[j], TextRange) else None)


def _FormatText(tr):
    # Set the text color to black
    tr.CharacterFormat.TextColor = Color.get_Black()
    # Set the text underline style to none
    tr.CharacterFormat.UnderlineStyle = UnderlineStyle.none
```

---

# Spire.Doc Python Hyperlink Formatting
## Set different hyperlink formats in Word documents
```python
# Add a paragraph and append a hyperlink to the paragraph
para1 = section.AddParagraph()
para1.AppendText("Regular Link: ")
# Format the hyperlink with default color and underline style
txtRange1 = para1.AppendHyperlink(
    "www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink)
txtRange1.CharacterFormat.FontName = "Times New Roman"
txtRange1.CharacterFormat.FontSize = 12
blankPara1 = section.AddParagraph()

# Add a paragraph and append a hyperlink to the paragraph
para2 = section.AddParagraph()
para2.AppendText("Change Color: ")
# Format the hyperlink with red color and underline style
txtRange2 = para2.AppendHyperlink(
    "www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink)
txtRange2.CharacterFormat.FontName = "Times New Roman"
txtRange2.CharacterFormat.FontSize = 12
txtRange2.CharacterFormat.TextColor = Color.get_Red()
blankPara2 = section.AddParagraph()

# Add a paragraph and append a hyperlink to the paragraph
para3 = section.AddParagraph()
para3.AppendText("Remove Underline: ")
# Format the hyperlink with red color and no underline style
txtRange3 = para3.AppendHyperlink(
    "www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink)
txtRange3.CharacterFormat.FontName = "Times New Roman"
txtRange3.CharacterFormat.FontSize = 12
txtRange3.CharacterFormat.UnderlineStyle = UnderlineStyle.none
```

---

# Spire.Doc Python Decryption
## Decrypt a password-protected Word document
```python
outputFile = "Decrypt.docx"
inputFile = "./Data/TemplateWithPassword.docx"

# Create word document
document = Document()
document.LoadFromFile(inputFile, FileFormat.Docx, "E-iceblue")

# Save as doc file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
```

---

# Spire.Doc Python Document Encryption
## Encrypt a Word document with a password
```python
# Create word document
document = Document()

# Load Word document
document.LoadFromFile(inputFile)

# Encrypt document with password
document.Encrypt("E-iceblue")

# Save as docx file
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()
```

---

# Spire.Doc Python Security
## Lock specified sections in a Word document
```python
# Create Word document.
document = Document()

# Add new sections.
s1 = document.AddSection()
s2 = document.AddSection()

# Protect the document with AllowOnlyFormFields protection type.
document.Protect(ProtectionType.AllowOnlyFormFields, "123")

# Unprotect section 2
s2.ProtectForm = False
```

---

# spire.doc python document security
## remove editable ranges from Word document
```python
# Find "PermissionStart" and "PermissionEnd" tags and remove them
for k in range(document.Sections.Count):
    section = document.Sections.get_Item(k)
    for j in range(section.Body.Paragraphs.Count):
        paragraph = section.Body.Paragraphs.get_Item(j)
        i = 0
        while i < paragraph.ChildObjects.Count:
            obj = paragraph.ChildObjects[i]
            if isinstance(obj, PermissionStart) or isinstance(obj, PermissionEnd):
                paragraph.ChildObjects.Remove(obj)
            else:
                i += 1
```

---

# Spire.Doc Python Security
## Remove read-only restriction from Word document
```python
# Remove ReadOnly Restriction.
doc.Protect(ProtectionType.NoProtection)
```

---

# Spire.Doc Python Security
## Set editable ranges in a protected document
```python
# Protect whole document
document.Protect(ProtectionType.AllowOnlyReading, "password")
# Create tags for permission start and end
start = PermissionStart(document, "testID")
end = PermissionEnd(document, "testID")
# Add the start and end tags to allow the first paragraph to be edited
document.Sections[0].Paragraphs[0].ChildObjects.Insert(0, start)
document.Sections[0].Paragraphs[0].ChildObjects.Add(end)
```

---

# Spire.Doc Document Protection
## Demonstrates how to protect a Word document with a specified protection type
```python
# Create Word document.
document = Document()

# Protect the Word file.
document.Protect(ProtectionType.AllowOnlyReading, "123456")
```

---

# Word to PDF with Encryption
## Convert a Word document to an encrypted PDF file
```python
inputFile = "./Data/Template_Docx_2.docx"
outputFile = "WordToPdfEncrypt.pdf"

#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

#Create an instance of ToPdfParameterList.
toPdf = ToPdfParameterList()

#Set the user password for the resulted PDF file.
toPdf.PdfSecurity.Encrypt("e-iceblue")

#Save to file.
document.SaveToFile(outputFile, toPdf)
document.Close()
```

---

# spire.doc python TC field
## add Table of Contents Entry field to Word document
```python
# Create Word document.
document = Document()

# Add a new section.
section = document.AddSection()

# Add a new paragraph.
paragraph = section.AddParagraph()

# Add TC field in the paragraph
field = paragraph.AppendField("TC", FieldType.FieldTOCEntry)
field.Code = """TC """ + "\"Entry Text\"" + " \\f" + " t"
```

---

# Spire.Doc Equation Conversion
## Convert equation fields to OfficeMath objects
```python
# Get the first paragraph of the first section in the document
paragraph = document.Sections.get_Item(0).Paragraphs.get_Item(0)

# Iterate through the child objects of the paragraph
i = 0
while i < paragraph.ChildObjects.Count:
    # Get the current document object
    documentObject = paragraph.ChildObjects[i]

    # Check if the document object is a field of type Equation
    if isinstance(documentObject,
                  Field) and documentObject.Type == FieldType.FieldEquation:
        # Convert the field to an OfficeMath object
        officeMath = OfficeMath.FromEqField(documentObject)

        # If conversion is successful, replace the field with the OfficeMath object
        if officeMath is not None:
            paragraph.ChildObjects.Remove(documentObject)
            paragraph.ChildObjects.Insert(i, officeMath)
    i += 1
```

---

# spire.doc python field conversion
## convert form fields to body text in document
```python
# Create the source document
sourceDocument = Document()

# Traverse FormFields
for j in range(sourceDocument.Sections[0].Body.FormFields.Count):
    field = sourceDocument.Sections[0].Body.FormFields.get_Item(j)
    # Find FieldFormTextInput type field
    if field.Type == FieldType.FieldFormTextInput:
        # Get the paragraph
        paragraph = field.OwnerParagraph

        # Define variables
        startIndex = 0
        endIndex = 0

        # Create a new TextRange
        textRange = TextRange(sourceDocument)

        # Set text for textRange
        textRange.Text = paragraph.Text

        # Traverse DocumentObjectS of field paragraph
        for k in range(paragraph.ChildObjects.Count):
            obj = paragraph.ChildObjects.get_Item(k)
            # If its DocumentObjectType is BookmarkStart
            if obj.DocumentObjectType == DocumentObjectType.BookmarkStart:
                # Get the index
                startIndex = paragraph.ChildObjects.IndexOf(obj)
            # If its DocumentObjectType is BookmarkEnd
            if obj.DocumentObjectType == DocumentObjectType.BookmarkEnd:
                # Get the index
                endIndex = paragraph.ChildObjects.IndexOf(obj)
        # Remove ChildObjects
        for i in range(endIndex, startIndex, -1):
            # If it is TextFormField
            if isinstance(paragraph.ChildObjects[i], TextFormField):
                textFormField = paragraph.ChildObjects[i] if isinstance(
                    paragraph.ChildObjects[i], TextFormField) else None

                # Remove the field object
                paragraph.ChildObjects.Remove(textFormField)
            else:
                paragraph.ChildObjects.RemoveAt(i)
        # Insert the new TextRange
        paragraph.ChildObjects.Insert(startIndex, textRange)
        break
```

---

# spire.doc python field conversion
## convert document fields to text
```python
# Get all fields in document
fields = document.Fields
count = fields.Count

for i in range(0, count):
    field = fields[0]
    s = field.FieldText
    index = field.OwnerParagraph.ChildObjects.IndexOf(field)
    textRange = TextRange(document)
    textRange.Text = s
    textRange.CharacterFormat.FontSize = 24

    field.OwnerParagraph.ChildObjects.Insert(index, textRange)
    field.OwnerParagraph.ChildObjects.Remove(field)
```

---

# Spire.Doc Python Field Conversion
## Convert IF fields to text in Word document
```python
# Get all fields in document
fields = document.Fields

for i in range(fields.Count):
    field = fields[i]
    if field.Type == FieldType.FieldIf:
        original = field if isinstance(field, TextRange) else None
        # Get field text
        text = field.FieldText
        # Create a new textRange and set its format
        textRange = TextRange(document)
        textRange.Text = text
        textRange.CharacterFormat.FontName = original.CharacterFormat.FontName
        textRange.CharacterFormat.FontSize = original.CharacterFormat.FontSize

        par = field.OwnerParagraph
        # Get the index of the if field
        index = par.ChildObjects.IndexOf(field)
        # Remove if field via index
        par.ChildObjects.RemoveAt(index)
        # Insert field text at the position of if field
        par.ChildObjects.Insert(index, textRange)
```

---

# Spire.Doc Python Cross-Reference
## Create a cross-reference field linked to a bookmark in a Word document
```python
# Create Word document.
document = Document()

# Add a new section.
section = document.AddSection()

# Create a bookmark.
paragraph = section.AddParagraph()
paragraph.AppendBookmarkStart("MyBookmark")
paragraph.AppendText("Text inside a bookmark")
paragraph.AppendBookmarkEnd("MyBookmark")

# Create a cross-reference field, and link it to bookmark.
field = Field(document)
field.Type = FieldType.FieldRef
field.Code = """REF MyBookmark \\p \\h"""

# Insert field to paragraph.
paragraph = section.AddParagraph()
paragraph.AppendText("For more information, see ")
paragraph.ChildObjects.Add(field)

# Insert FieldSeparator object.
fieldSeparator = FieldMark(document, FieldMarkType.FieldSeparator)
paragraph.ChildObjects.Add(fieldSeparator)

# Set display text of the field.
tr = TextRange(document)
tr.Text = "above"
paragraph.ChildObjects.Add(tr)

# Insert FieldEnd object to mark the end of the field.
fieldEnd = FieldMark(document, FieldMarkType.FieldEnd)
paragraph.ChildObjects.Add(fieldEnd)
```

---

# Spire.Doc Python Form Fields
## Create different types of form fields in a Word document
```python
# Add table for form fields
table = section.AddTable()
table.DefaultColumnsNumber = 2
table.DefaultRowHeight = 20

# Create a row for field group label
row = table.AddRow(False)
row.Cells[0].CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(0xFF, 0x00, 0x71, 0xb6)
row.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle

# Add label for field group
cellParagraph = row.Cells[0].AddParagraph()
cellParagraph.AppendText(tempNameStr)

# Create rows for form fields
for fieldNode in fieldNodes:
    # Create a row for field
    fieldRow = table.AddRow(False)
    
    # Add field label
    fieldRow.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle
    labelParagraph = fieldRow.Cells[0].AddParagraph()
    labelParagraph.AppendText(fieldNode.get("label", ""))
    
    # Add form field based on type
    fieldRow.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle
    fieldParagraph = fieldRow.Cells[1].AddParagraph()
    fieldId = fieldNode.get("id", "")
    
    if fieldNode.get("type", "") == "text":
        # Add text input field
        fieldFormTextInput = fieldParagraph.AppendField(fieldId, FieldType.FieldFormTextInput)
        field = fieldFormTextInput if isinstance(fieldFormTextInput, TextFormField) else None
        field.DefaultText = ""
        field.Text = ""
    elif fieldNode.get("type", "") == "list":
        # Add dropdown field
        fieldFormDropDown = fieldParagraph.AppendField(fieldId, FieldType.FieldFormDropDown)
        fieldList = fieldFormDropDown if isinstance(fieldFormDropDown, DropDownFormField) else None
        itemNodes = fieldNode.findall(".//item")
        for itemNode in itemNodes:
            fieldList.DropDownItems.Add(itemNode.text)
    elif fieldNode.get("type", "") == "checkbox":
        # Add checkbox field
        fieldParagraph.AppendField(fieldId, FieldType.FieldFormCheckBox)

# Merge field group row
table.ApplyHorizontalMerge(row.GetRowIndex(), 0, 1)
```

---

# Spire.Doc Python IF Field Creation
## Create an IF field in a Word document that displays different messages based on a condition

```python
def _CreateIfField(document, paragraph):
    ifField = IfField(document)
    ifField.Type = FieldType.FieldIf
    ifField.Code = "IF "

    paragraph.Items.Add(ifField)
    paragraph.AppendField("Count", FieldType.FieldMergeField)
    paragraph.AppendText(" > ")
    paragraph.AppendText("\"100\" ")
    paragraph.AppendText("\"Thanks\" ")
    paragraph.AppendText("\"The minimum order is 100 units\"")

    end = document.CreateParagraphItem(ParagraphItemType.FieldMark)
    (end if isinstance(end, FieldMark) else None).Type = FieldMarkType.FieldEnd
    paragraph.Items.Add(end)
    ifField.End = end if isinstance(end, FieldMark) else None
```

---

# Spire.Doc Python Nested Fields
## Create nested IF fields in a Word document
```python
# Get the first section
section = document.Sections[0]

paragraph = section.AddParagraph()

# Create an IF field
ifField = IfField(document)
ifField.Type = FieldType.FieldIf
ifField.Code = "IF "
paragraph.Items.Add(ifField)

# Create the embedded IF field
ifField2 = IfField(document)
ifField2.Type = FieldType.FieldIf
ifField2.Code = "IF "
paragraph.ChildObjects.Add(ifField2)
paragraph.Items.Add(ifField2)
paragraph.AppendText("\"200\" < \"50\"   \"200\" \"50\" ")
embeddedEnd = document.CreateParagraphItem(ParagraphItemType.FieldMark)
(embeddedEnd if isinstance(embeddedEnd, FieldMark)
 else None).Type = FieldMarkType.FieldEnd
paragraph.Items.Add(embeddedEnd)
ifField2.End = embeddedEnd if isinstance(embeddedEnd, FieldMark) else None

paragraph.AppendText(" > ")
paragraph.AppendText("\"100\" ")
paragraph.AppendText("\"Thanks\" ")
paragraph.AppendText("\"The minimum order is 100 units\"")
end = document.CreateParagraphItem(ParagraphItemType.FieldMark)
(end if isinstance(end, FieldMark) else None).Type = FieldMarkType.FieldEnd
paragraph.Items.Add(end)
ifField.End = end if isinstance(end, FieldMark) else None

# Update all fields in the document.
document.IsUpdateFields = True
```

---

# Spire.Doc Python Form Field Filling
## Fill form fields in a Word document with different field types
```python
# Fill form fields in a document
for k in range(document.Sections[0].Body.FormFields.Count):
    field = document.Sections[0].Body.FormFields.get_Item(k)
    
    # Handle text input fields
    if field.Type == FieldType.FieldFormTextInput:
        field.Text = "Text value"
        
    # Handle dropdown fields
    elif field.Type == FieldType.FieldFormDropDown:
        combox = field if isinstance(field, DropDownFormField) else None
        for i in range(combox.DropDownItems.Count):
            if combox.DropDownItems[i].Text == "Selected option":
                combox.DropDownSelectedIndex = i
                break
            if field.Name == "country" and combox.DropDownItems[i].Text == "Others":
                combox.DropDownSelectedIndex = i
                
    # Handle checkbox fields
    elif field.Type == FieldType.FieldFormCheckBox:
        if True:  # Condition to check/uncheck
            checkBox = field if isinstance(field, CheckBoxFormField) else None
            checkBox.Checked = True
```

---

# Spire.Doc Python Form Fields
## Modify form field properties in Word documents
```python
# Get the first section
section = document.Sections[0]

# Get FormField by index
formField = section.Body.FormFields[1]

if formField.Type == FieldType.FieldFormTextInput:
    formField.Text = "My name is " + formField.Name
    formField.CharacterFormat.TextColor = Color.get_Red()
    formField.CharacterFormat.Italic = True
```

---

# Spire.Doc Python Get Field Text
## Extract text from fields in a Word document
```python
# Get all fields in document
fields = document.Fields

for i in range(fields.Count):
    field = fields.get_Item(i)
    # Get field text
    fieldText = field.FieldText
```

---

# Spire.Doc Python Form Field
## Get form field by name from Word document
```python
# Create a Word document
document = Document()

# Get the first section
section = document.Sections[0]

# Get form field by name
formField = section.Body.FormFields["email"]
```

---

# Spire.Doc form fields collection
## Get form fields collection from document section
```python
# Get the first section
section = document.Sections[0]

# Get form fields from the section body
formFields = section.Body.FormFields

# Get the count of form fields
formFields.Count
```

---

# Spire.Doc Get Merge Field Names
## Extract merge field names from a Word document
```python
# Open a Word document
document = Document()
document.LoadFromFile(inputFile)

# Get merge field names
fieldNames = document.MailMerge.GetMergeFieldNames()

document.Close()
```

---

# spire.doc python address block field
## insert address block field in word document
```python
# Get the first section
section = document.Sections[0]

par = section.AddParagraph()

# Add address block in the paragraph
field = par.AppendField("ADDRESSBLOCK", FieldType.FieldAddressBlock)

# Set field code
field.Code = "ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\""
```

---

# Spire.Doc Python Advance Field
## Insert and configure an advance field in a Word document
```python
# Create a Word document
document = Document()

# Get the first section
section = document.Sections[0]

par = section.AddParagraph()

# Add advance field
field = par.AppendField("Field", FieldType.FieldAdvance)

# Add field code
field.Code = "ADVANCE \\d 10 \\l 10 \\r 10 \\u 0 \\x 100 \\y 100 "

# Update field
document.IsUpdateFields = True
```

---

# Spire.Doc Python Merge Field
## Insert a merge field in a Word document
```python
document = Document()
section = document.Sections[0]
par = section.AddParagraph()
field = MergeField(par.AppendField(
    "MyFieldName", FieldType.FieldMergeField))
```

---

# Insert None Field in Word Document
## Demonstrate how to insert a none field into a Word document
```python
# Get the first section
section = document.Sections[0]

par = section.AddParagraph()

# Add a none field
field = par.AppendField('', FieldType.FieldNone)
```

---

# Spire.Doc Python Page Reference Field
## Insert a page reference field in a Word document
```python
# Get the first section
section = document.LastSection

par = section.AddParagraph()

# Add page ref field
field = par.AppendField("pageRef", FieldType.FieldPageRef)

# Set field code
field.Code = "PAGEREF  bookmark1 \\# \"0\" \\* Arabic  \\* MERGEFORMAT"

# Update field
document.IsUpdateFields = True
```

---

# spire.doc python remove custom properties
## remove all custom property fields from a Word document
```python
# Get custom document properties object.
cdp = document.CustomDocumentProperties

# Remove all custom property fields in the document.
i = 0
while i < cdp.Count:
    cdp.Remove(cdp[i].Name)

document.IsUpdateFields = True
```

---

# Spire.Doc Python Field Removal
## Remove a field from a Word document
```python
#Get the first field
field = document.Fields[0]
#Get the paragraph of the field
par = field.OwnerParagraph
#Get the index of the field
index = par.ChildObjects.IndexOf(field)
#Remove field via index
par.ChildObjects.RemoveAt(index)
```

---

# spire.doc python field replacement
## replace text with merge field
```python
#Find the text that will be replaced
ts = document.FindString("Test", True, True)
tr = ts.GetAsOneRange()

#Get the paragraph
par = tr.OwnerParagraph

#Get the index of the text in the paragraph
index = par.ChildObjects.IndexOf(tr)

#Create a new field
field = MergeField(document)
field.FieldName = "MergeField"

#Insert field at specific position
par.ChildObjects.Insert(index, field)

#Remove the text
par.ChildObjects.Remove(tr)
```

---

# Spire.Doc Python Field Locale
## Set locale for field in document
```python
#Get the first section
section = document.Sections[0]

par = section.AddParagraph()

#Add a date field
field = par.AppendField("DocDate", FieldType.FieldDate)

#Set the LocaleId for the textrange
( field.OwnerParagraph.ChildObjects[0] if isinstance(field.OwnerParagraph.ChildObjects[0], TextRange) else None).CharacterFormat.LocaleIdASCII = 1049

field.FieldText = "2019-10-10"
#Update field
document.IsUpdateFields = True
```

---

# spire.doc python field update
## Update fields in a Word document
```python
# Update fields
document.IsUpdateFields = True
```

---

# Spire.Doc Python TOC Style
## Change Table of Contents style in Word document
```python
# Create document
doc = Document()

# Define a TOC style
tocStyle = Style.CreateBuiltinStyle(BuiltinStyle.Toc1, doc) if isinstance(Style.CreateBuiltinStyle(BuiltinStyle.Toc1, doc), ParagraphStyle) else None
tocStyle.CharacterFormat.FontName = "Aleo"
tocStyle.CharacterFormat.FontSize = 15
tocStyle.CharacterFormat.TextColor = Color.get_CadetBlue()
doc.Styles.Add(tocStyle)

# Loop through sections
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    # Loop through content of section
    for j in range(section.Body.ChildObjects.Count):
        obj = section.Body.ChildObjects.get_Item(j)
        # Find the structure document tag
        if isinstance(obj, StructureDocumentTag):
            tag = obj if isinstance(obj, StructureDocumentTag) else None
            # Find the paragraph where the TOC1 locates
            for k in range(tag.ChildObjects.Count):
                cObj = tag.ChildObjects.get_Item(k)
                if isinstance(cObj, Paragraph):
                    para = cObj if isinstance(cObj, Paragraph) else None
                    if para.StyleName == "TOC1":
                        # Apply the new style for TOC1 paragraph
                        para.ApplyStyle(tocStyle.Name)
```

---

# Spire.Doc Python Table of Contents
## Change TOC tab style in Word document
```python
#Loop through sections
for k in range(doc.Sections.Count):
    section = doc.Sections.get_Item(k)
    #Loop through content of section
    for j in range(section.Body.ChildObjects.Count):
        obj = section.Body.ChildObjects.get_Item(j)
        #Find the structure document tag
        if isinstance(obj, StructureDocumentTag):
            tag = obj if isinstance(obj, StructureDocumentTag) else None
            #Find the paragraph where the TOC1 locates
            for m in range(tag.ChildObjects.Count):
                cObj = tag.ChildObjects.get_Item(m)
                if isinstance(cObj, Paragraph):
                    para = cObj if isinstance(cObj, Paragraph) else None
                    if para.StyleName == "TOC2":
                        #Set the tab style of paragraph
                        for n in range(para.Format.Tabs.Count):
                            tab = para.Format.Tabs.get_Item(n)
                            tab.Position = tab.Position + 20
                            tab.TabLeader = TabLeader.NoLeader
```

---

# Spire.Doc Python Table of Contents
## Create a table of contents with default settings in a Word document
```python
doc = Document()
section = doc.AddSection()
para = section.AddParagraph()
#Create table of content with default switches(\o "1-3" \h \z)
para.AppendTOC(1, 3)

#Create paragraph and set the head level
para1 = section.AddParagraph()
para1.AppendText("Ornithogalum")
#Apply the Heading1 style
para1.ApplyStyle(BuiltinStyle.Heading1)

para2 = section.AddParagraph()
para2.AppendText("Rosa")
#Apply the Heading2 style
para2.ApplyStyle(BuiltinStyle.Heading2)

para3 = section.AddParagraph()
para3.AppendText("Hyacinth")
#Apply the Heading3 style
para3.ApplyStyle(BuiltinStyle.Heading3)

#Update TOC
doc.UpdateTableOfContents()
```

---

# Spire.Doc Python Table of Contents Customization
## Customize and update a table of contents in a Word document
```python
#Create a document
doc = Document()
#Add a section
section = doc.AddSection()
#Customize table of contents with switches
toc = TableOfContent(doc, "{\\o \"1-3\" \\n 1-1}")
para = section.AddParagraph()
para.Items.Add(toc)
para.AppendFieldMark(FieldMarkType.FieldSeparator)
para.AppendText("TOC")
para.AppendFieldMark(FieldMarkType.FieldEnd)
doc.TOC = toc

#Create content with different heading styles
para1 = section.AddParagraph()
para1.AppendText("Heading 1 Text")
#Apply the Heading1 style
para1.ApplyStyle(BuiltinStyle.Heading1)

para2 = section.AddParagraph()
para2.AppendText("Heading 2 Text")
#Apply the Heading2 style
para2.ApplyStyle(BuiltinStyle.Heading2)

para3 = section.AddParagraph()
para3.AppendText("Heading 3 Text")
#Apply the Heading3 style
para3.ApplyStyle(BuiltinStyle.Heading3)

#Update TOC
doc.UpdateTableOfContents()
```

---

# spire.doc python table of content
## remove table of content from document
```python
#Get the first body from the first section
body = document.Sections[0].Body

#Remove TOC from first body
regexStr = "TOC\\w+"
i = 0
while i < body.Paragraphs.Count:
    if re.match(regexStr,body.Paragraphs[i].StyleName):
        body.Paragraphs.RemoveAt(i)
        i -= 1
    i += 1
```

---

# Delete Table From TextBox
## Delete a table from a textbox in a Word document
```python
#Get the first textbox
textbox = doc.TextBoxes[0]

#Remove the first table from the textbox
textbox.Body.Tables.RemoveAt(0)
```

---

# spire.doc python text extraction
## extract text from textboxes in a Word document
```python
#Verify whether the document contains a textbox or not.
if document.TextBoxes.Count > 0:
    #Traverse the document.
    for i in range(document.Sections.Count):
        section = document.Sections.get_Item(i)
        for j in range(section.Paragraphs.Count):
            p = section.Paragraphs.get_Item(j)
            for k in range(p.ChildObjects.Count):
                obj = p.ChildObjects.get_Item(k)
                if obj.DocumentObjectType == DocumentObjectType.TextBox:
                    textbox = obj if isinstance(obj, TextBox) else None
                    for x in range(textbox.ChildObjects.Count):
                        objt = textbox.ChildObjects.get_Item(x)
                        #Extract text from paragraph in TextBox.
                        if objt.DocumentObjectType == DocumentObjectType.Paragraph:
                            # Get text from paragraph
                            text = (objt if isinstance(objt, Paragraph) else None).Text
                        #Extract text from Table in TextBox.
                        if objt.DocumentObjectType == DocumentObjectType.Table:
                            table = objt if isinstance(objt, Table) else None
                            for i in range(table.Rows.Count):
                                row = table.Rows[i]
                                for j in range(row.Cells.Count):
                                    cell = row.Cells[j]
                                    for k in range(cell.Paragraphs.Count):
                                        paragraph = cell.Paragraphs.get_Item(k)
                                        # Get text from table cell
                                        text = paragraph.Text
```

---

# Insert Image into Textbox
## This code demonstrates how to insert an image into a textbox in a Word document
```python
#Create a new document
doc = Document()
section = doc.AddSection()
paragraph = section.AddParagraph()

#Append a textbox to paragraph
tb = paragraph.AppendTextBox(220, 220)

#Set the position of the textbox
tb.Format.HorizontalOrigin = HorizontalOrigin.Page
tb.Format.HorizontalPosition = 50
tb.Format.VerticalOrigin = VerticalOrigin.Page
tb.Format.VerticalPosition = 50

#Set the fill effect of textbox as picture
tb.Format.FillEfects.Type = BackgroundType.Picture

#Fill the textbox with a picture
tb.Format.FillEfects.SetPicture("./Data/Spire.Doc.png")
```

---

# Spire.Doc Python Textbox Table
## Insert a table into a textbox in a Word document
```python
#Create a new document
doc = Document()

#Add a section
section = doc.AddSection()

#Add a paragraph to the section
paragraph = section.AddParagraph()

#Add a textbox to the paragraph
textbox = paragraph.AppendTextBox(300, 100)

#Set the position of the textbox
textbox.Format.HorizontalOrigin = HorizontalOrigin.Page
textbox.Format.HorizontalPosition = 140
textbox.Format.VerticalOrigin = VerticalOrigin.Page
textbox.Format.VerticalPosition = 50

#Add text to the textbox
textboxParagraph = textbox.Body.AddParagraph()
textboxRange = textboxParagraph.AppendText("Table 1")
textboxRange.CharacterFormat.FontName = "Arial"

#Insert table to the textbox
table = textbox.Body.AddTable(True)

#Specify the number of rows and columns of the table
table.ResetCells(4, 4)
data = [["Name", "Age", "Gender", "ID"], ["John", "28", "Male", "0023"], ["Steve", "30", "Male", "0024"], ["Lucy", "26", "female", "0025"]]

#Add data to the table 
for i in range(0, 4):
    for j in range(0, 4):
        tableRange = table.Rows[i].Cells[j].AddParagraph().AppendText(data[i][j])
        tableRange.CharacterFormat.FontName = "Arial"

#Apply style to the table
table.ApplyStyle(DefaultTableStyle.TableColorful2)
```

---

# Spire.Doc Python Textbox
## Create a textbox with locked aspect ratio in a Word document
```python
# Create a new instance of Document
document = Document()

# Add a new section to the document
section = document.AddSection()

# Add a paragraph to the section
paragraph = section.AddParagraph()

# Append a textbox to the paragraph and get a reference to it
textBox1 = paragraph.AppendTextBox(240, 35)

# Configure the horizontal alignment, line color, and line style of the textbox
textBox1.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left
textBox1.Format.LineColor = Color.get_Black()
textBox1.Format.LineStyle = TextBoxLineStyle.Simple

# Lock the aspect ratio of the textbox
textBox1.AspectRatioLocked = True

# Add a paragraph to the body of the textbox and get a reference to it
para = textBox1.Body.AddParagraph()

# Add text to the paragraph
txtrg = para.AppendText("Textbox 1 in the document")
txtrg.CharacterFormat.FontName = "Lucida Sans Unicode"
txtrg.CharacterFormat.FontSize = 14
txtrg.CharacterFormat.TextColor = Color.get_Black()
para.Format.HorizontalAlignment = HorizontalAlignment.Center
```

---

# spire.doc python textbox
## read table from textbox
```python
#Get the first textbox
textbox = doc.TextBoxes[0]
#Get the first table in the textbox
table = textbox.Body.Tables[0] if isinstance(textbox.Body.Tables[0], Table) else None
```

---

# spire.doc python textbox
## remove textbox from document
```python
#Remove the first text box
doc.TextBoxes.RemoveAt(0)

#Clear all the text boxes
#Doc.TextBoxes.Clear()
```

---

# spire.doc python textbox
## create textboxes with different formatting styles in a Word document
```python
#Create a Word document and a section.
document = Document()
section = document.AddSection()
paragraph = section.Paragraphs[0] if section.Paragraphs.Count > 0 else section.AddParagraph()
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()

#Insert and format the first textbox.
textBox1 = paragraph.AppendTextBox(240, 35)
textBox1.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left
textBox1.Format.LineColor = Color.get_Gray()
textBox1.Format.LineStyle = TextBoxLineStyle.Simple
textBox1.Format.FillColor = Color.get_DarkSeaGreen()
para = textBox1.Body.AddParagraph()
txtrg = para.AppendText("Textbox 1 in the document")
txtrg.CharacterFormat.FontName = "Lucida Sans Unicode"
txtrg.CharacterFormat.FontSize = 14
txtrg.CharacterFormat.TextColor = Color.get_White()
para.Format.HorizontalAlignment = HorizontalAlignment.Center

#Insert and format the second textbox.
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()
textBox2 = paragraph.AppendTextBox(240, 35)
textBox2.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left
textBox2.Format.LineColor = Color.get_Tomato()
textBox2.Format.LineStyle = TextBoxLineStyle.ThinThick
textBox2.Format.FillColor = Color.get_Blue()
textBox2.Format.LineDashing = LineDashing.Dot
para = textBox2.Body.AddParagraph()
txtrg = para.AppendText("Textbox 2 in the document")
txtrg.CharacterFormat.FontName = "Lucida Sans Unicode"
txtrg.CharacterFormat.FontSize = 14
txtrg.CharacterFormat.TextColor = Color.get_Pink()
para.Format.HorizontalAlignment = HorizontalAlignment.Center

#Insert and format the third textbox.
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()
paragraph = section.AddParagraph()
textBox3 = paragraph.AppendTextBox(240, 35)
textBox3.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left
textBox3.Format.LineColor = Color.get_Violet()
textBox3.Format.LineStyle = TextBoxLineStyle.Triple
textBox3.Format.FillColor = Color.get_Pink()
textBox3.Format.LineDashing = LineDashing.DashDotDot
para = textBox3.Body.AddParagraph()
txtrg = para.AppendText("Textbox 3 in the document")
txtrg.CharacterFormat.FontName = "Lucida Sans Unicode"
txtrg.CharacterFormat.FontSize = 14
txtrg.CharacterFormat.TextColor = Color.get_Tomato()
para.Format.HorizontalAlignment = HorizontalAlignment.Center
```

---

# Spire.Doc Python TextBox Formatting
## Create and format a text box in a Word document
```python
#Create a new document
doc = Document()
sec = doc.AddSection()

#Add a text box and append sample text
TB = doc.Sections[0].AddParagraph().AppendTextBox(310, 90)
para = TB.Body.AddParagraph()
TR = para.AppendText("Using Spire.Doc, developers will find " + "a simple and effective method to endow their applications with rich MS Word features. ")
TR.CharacterFormat.FontName = "Cambria "
TR.CharacterFormat.FontSize = 13

#Set exact position for the text box
TB.Format.HorizontalOrigin = HorizontalOrigin.Page
TB.Format.HorizontalPosition = 120
TB.Format.VerticalOrigin = VerticalOrigin.Page
TB.Format.VerticalPosition = 100

#Set line style for the text box
TB.Format.LineStyle = TextBoxLineStyle.Double
TB.Format.LineColor = Color.get_CornflowerBlue()
TB.Format.LineDashing = LineDashing.Solid
TB.Format.LineWidth = 5

#Set internal margin for the text box
TB.Format.InternalMargin.Top = 15
TB.Format.InternalMargin.Bottom = 10
TB.Format.InternalMargin.Left = 12
TB.Format.InternalMargin.Right = 10
```

---

# Spire.Doc Python Image Watermark
## Add image watermark to Word document
```python
# Create a document
document = Document()

# Insert the image watermark
picture = PictureWatermark()
picture.SetPicture("./Data/ImageWatermark.png")
picture.Scaling = 250
picture.IsWashout = False
document.Watermark = picture
```

---

# Remove Image Watermark from Word Document
## This code demonstrates how to remove image watermarks from a Word document using Spire.Doc for Python
```python
#Set the watermark as null to remove the text and image watermark.
document.Watermark = None
```

---

# spire.doc python watermark
## remove text watermark from Word document
```python
#Set the watermark as null to remove the text and image watermark.
document.Watermark = None
```

---

# Spire.Doc Text Watermark
## Create and apply a text watermark to a Word document
```python
# Insert text watermark
txtWatermark = TextWatermark()
txtWatermark.Text = "E-iceblue"
txtWatermark.FontSize = 95
txtWatermark.Color = Color.get_Blue()
txtWatermark.Layout = WatermarkLayout.Diagonal
document.Watermark = txtWatermark
```

---

# spire.doc python OLE extraction
## extract OLE objects from Word document and save as different file formats
```python
# Create document and load file from disk
doc = Document()
doc.LoadFromFile(inputFile)

# Traverse through all sections of the word document    
for k in range(doc.Sections.Count):
    sec = doc.Sections.get_Item(k)
    # Traverse through all Child Objects in the body of each section
    for j in range(sec.Body.ChildObjects.Count):
        obj = sec.Body.ChildObjects.get_Item(j)
        # find the paragraph
        if isinstance(obj, Paragraph):
            par = obj if isinstance(obj, Paragraph) else None
            for m in range(par.ChildObjects.Count):
                o = par.ChildObjects.get_Item(m)
                # check whether the object is OLE
                if o.DocumentObjectType == DocumentObjectType.OleObject:
                    Ole = o if isinstance(o, DocOleObject) else None
                    s = Ole.ObjectType
                    # check whether the object type is "Acrobat.Document.11"
                    if s == "AcroExch.Document.DC":
                        # write the data of OLE into file
                        fp = open(outputFile_pdf,"wb")
                        fp.write(Ole.NativeData)
                        fp.close()

                    # check whether the object type is "Excel.Sheet.8"
                    elif s == "Excel.Sheet.8":
                        fp = open(outputFile_xls,"wb")
                        fp.write(Ole.NativeData)
                        fp.close()
                    # check whether the object type is "PowerPoint.Show.12"
                    elif s == "PowerPoint.Show.12":
                        fp = open(outputFile_pptx,"wb")
                        fp.write(Ole.NativeData)
                        fp.close()
doc.Close()
```

---

# Spire.Doc Python OLE
## Insert OLE object into Word document
```python
#create a document
doc = Document()

#add a section
sec = doc.AddSection()

#add a paragraph
par = sec.AddParagraph()

#load the image
picture = DocPicture(doc)
picture.LoadImage("./Data/Excel.png")

#insert the OLE
obj = par.AppendOleObject("./Data/example.xlsx", picture, OleObjectType.ExcelWorksheet)
```

---

# spire.doc python OLE object
## insert OLE object as icon via stream
```python
#Create word document
doc = Document()
#add a section
sec = doc.AddSection()
#add a paragraph
par = sec.AddParagraph()

#ole stream
stream = Stream(inputFile)

#load the image
picture = DocPicture(doc)
picture.LoadImage(inputFile_I)

#insert the OLE from stream
obj = par.AppendOleObject(stream, picture, "zip")

#display as icon
obj.DisplayAsIcon = True
```

---

# spire.doc python checkbox content control
## add checkbox content control to word document
```python
#Create a document
document = Document()

#Add a new section.
section = document.AddSection()

#Add a paragraph
paragraph = section.AddParagraph()

#Create StructureDocumentTagInline for document
sdt = StructureDocumentTagInline(document)

#Add sdt in paragraph
paragraph.ChildObjects.Add(sdt)

#Specify the type
sdt.SDTProperties.SDTType = SdtType.CheckBox

#Set properties for control
scb = SdtCheckBox()
sdt.SDTProperties.ControlProperties = scb

#Add textRange format
tr = TextRange(document)
tr.CharacterFormat.FontName = "MS Gothic"
tr.CharacterFormat.FontSize = 12

#Add textRange to StructureDocumentTagInline
sdt.ChildObjects.Add(tr)

#Set checkBox as checked
scb.Checked = True
```

---

# Spire.Doc Python Content Controls
## Add various content controls to a Word document
```python
# Add Combo Box Content Control
sd = StructureDocumentTagInline(document)
paragraph.ChildObjects.Add(sd)
sd.SDTProperties.SDTType = SdtType.ComboBox
cb = SdtComboBox()
cb.ListItems.Add(SdtListItem("Spire.Doc"))
cb.ListItems.Add(SdtListItem("Spire.XLS"))
cb.ListItems.Add(SdtListItem("Spire.PDF"))
sd.SDTProperties.ControlProperties = cb
rt = TextRange(document)
rt.Text = cb.ListItems[0].DisplayText
sd.SDTContent.ChildObjects.Add(rt)
```

## Add Text Content Control
```python
# Add Text Content Control
sd = StructureDocumentTagInline(document)
paragraph.ChildObjects.Add(sd)
sd.SDTProperties.SDTType = SdtType.Text
text = SdtText(True)
text.IsMultiline = True
sd.SDTProperties.ControlProperties = text
rt = TextRange(document)
rt.Text = "Text"
sd.SDTContent.ChildObjects.Add(rt)
```

## Add Picture Content Control
```python
# Add Picture Content Control
sd = StructureDocumentTagInline(document)
paragraph.ChildObjects.Add(sd)
sd.SDTProperties.SDTType = SdtType.Picture
pic = DocPicture(document)
pic.Width = 10
pic.Height = 10
pic.LoadImage("./Data/logo.png")
sd.SDTContent.ChildObjects.Add(pic)
```

## Add Date Picker Content Control
```python
# Add Date Picker Content Control
sd = StructureDocumentTagInline(document)
paragraph.ChildObjects.Add(sd)
sd.SDTProperties.SDTType = SdtType.DatePicker
date = SdtDate()
date.CalendarType = CalendarType.Default
date.DateFormat = "yyyy.MM.dd"
date.FullDate = DateTime.get_Now()
sd.SDTProperties.ControlProperties = date
rt = TextRange(document)
rt.Text = "1990.02.08"
sd.SDTContent.ChildObjects.Add(rt)
```

## Add Drop-Down List Content Control
```python
# Add Drop-Down List Content Control
sd = StructureDocumentTagInline(document)
paragraph.ChildObjects.Add(sd)
sd.SDTProperties.SDTType = SdtType.DropDownList
sddl = SdtDropDownList()
sddl.ListItems.Add(SdtListItem("Harry"))
sddl.ListItems.Add(SdtListItem("Jerry"))
sd.SDTProperties.ControlProperties = sddl
rt = TextRange(document)
rt.Text = sddl.ListItems[0].DisplayText
sd.SDTContent.ChildObjects.Add(rt)
```

---

# spire.doc python content control
## add rich text content control to word document
```python
#Create a document
document = Document()

#Add a new section
section = document.AddSection()

#Add a paragraph
paragraph = section.AddParagraph()

#Append textRange for the paragraph
txtRange = paragraph.AppendText("The following example shows how to add RichText content control in a Word document. \n")

#Append textRange 
txtRange = paragraph.AppendText("Add RichText Content Control:  ")

#Set the font format
txtRange.CharacterFormat.Italic = True

#Create StructureDocumentTagInline for document
sdt = StructureDocumentTagInline(document)

#Add sdt in paragraph
paragraph.ChildObjects.Add(sdt)

#Specify the type
sdt.SDTProperties.SDTType = SdtType.RichText

#Set displaying text
text = SdtText(True)
text.IsMultiline = True
sdt.SDTProperties.ControlProperties = text

#Create a TextRange
rt = TextRange(document)
rt.Text = "Welcome to use "
rt.CharacterFormat.TextColor = Color.get_Green()
sdt.SDTContent.ChildObjects.Add(rt)
rt = TextRange(document)
rt.Text = "Spire.Doc"
rt.CharacterFormat.TextColor = Color.get_OrangeRed()
sdt.SDTContent.ChildObjects.Add(rt)
```

---

# spire.doc python combo box manipulation
## modify combo box items in structured document tags
```python
#Get the combo box from the file
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    for j in range(section.Body.ChildObjects.Count):
        bodyObj = section.Body.ChildObjects.get_Item(j)
        if bodyObj.DocumentObjectType == DocumentObjectType.StructureDocumentTag:
            #If SDTType is ComboBox
            tempObj = bodyObj if isinstance(bodyObj, StructureDocumentTag) else None
            if tempObj.SDTProperties.SDTType == SdtType.ComboBox:
                tempProperties = tempObj.SDTProperties.ControlProperties
                combo = tempProperties if isinstance(tempProperties, SdtComboBox) else None
                #Remove the second list item
                combo.ListItems.RemoveAt(1)
                #Add a new item
                item = SdtListItem("D", "D")
                combo.ListItems.Add(item)

                #If the value of list items is "D"
                for k in range(combo.ListItems.Count):
                    sdtItem = combo.ListItems.get_Item(k)
                    if locale.strcoll(sdtItem.Value, 'D') == 0:
                        #Select the item
                        combo.ListItems.SelectedValue = sdtItem
```

---

# spire.doc content control properties
## extract properties from structured document tags in Word documents
```python
class StructureTags:
    def __init__(self):
        #instance fields found by C# to Python Converter:
        self._m_tagInlines = None
        self._m_tags = None

    def get_tag_inlines(self):
        if self._m_tagInlines is None:
            self._m_tagInlines = []
        return self._m_tagInlines
    def set_tag_inlines(self, value):
        self._m_tagInlines = value
    def get_tags(self):
        if self._m_tags is None:
            self._m_tags = []
        return self._m_tags
    def set_tags(self, value):
        self._m_tags = value

def _GetAllTags(document):
    structureTags = StructureTags()
    for i in range(document.Sections.Count):
        section = document.Sections.get_Item(i)
        for j in range(section.Body.ChildObjects.Count):
            obj = section.Body.ChildObjects.get_Item(j)
            if obj.DocumentObjectType == DocumentObjectType.StructureDocumentTag:
                structureTags.get_tags().append( obj if isinstance(obj, StructureDocumentTag) else None)

            elif obj.DocumentObjectType == DocumentObjectType.Paragraph:
                tempPara = obj if isinstance(obj, Paragraph) else None
                for k in range(tempPara.ChildObjects.Count):
                    pobj = tempPara.ChildObjects.get_Item(k)
                    if pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline:
                        tempPobj = pobj if isinstance(pobj, StructureDocumentTagInline) else None
                        structureTags.get_tag_inlines().append(tempPobj)
            elif obj.DocumentObjectType == DocumentObjectType.Table:
                tempTable = obj if isinstance(obj, Table) else None
                for x in range(tempTable.Rows.Count):
                    row = tempTable.Rows.get_Item(x)
                    for g in range(row.Cells.Count):
                        cell = row.Cells.get_Item(g)
                        for z in range(cell.ChildObjects.Count):
                            cellChild = cell.ChildObjects.get_Item(z)
                            if cellChild.DocumentObjectType == DocumentObjectType.StructureDocumentTag:
                                structureTags.get_tags().append( cellChild if isinstance(cellChild, StructureDocumentTag) else None)
                            elif cellChild.DocumentObjectType == DocumentObjectType.Paragraph:
                                tempParagraph = cellChild if isinstance(cellChild, Paragraph) else None
                            for p in range(tempParagraph.ChildObjects.Count):
                                pobj = tempParagraph.ChildObjects.get_Item(p)
                                if pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline:
                                    structureTags.get_tag_inlines().append( pobj if isinstance(pobj, StructureDocumentTagInline) else None)
    return structureTags

#Get all structureTags in the Word document
structureTags = _GetAllTags(doc)
#Get all StructureDocumentTagInline objects
tagInlines = structureTags.get_tag_inlines()
strProperty = ''
strProperty += "Alias of contentControl" + "\t" + "ID  " + "\t" + "Tag     " + "\t" + "STDType" + "\r\n"
#Get properties of all tagInlines
for i, unusedItem in enumerate(tagInlines):
    alias = tagInlines[i].SDTProperties.Alias
    objId = tagInlines[i].SDTProperties.Id
    tag = tagInlines[i].SDTProperties.Tag
    STDType = str(tagInlines[i].SDTProperties.SDTType)
    strProperty += alias + ",\t" + str(objId) + ",\t" + tag + ",\t" + STDType + "\r\n"

#Get all StructureDocumentTag objects
tags = structureTags.get_tags()
#Get properties of all tags
for i, unusedItem in enumerate(tags):
    alias = tags[i].SDTProperties.Alias
    objId = tags[i].SDTProperties.Id
    tag = tags[i].SDTProperties.Tag
    STDType = str(tags[i].SDTProperties.SDTType)
    strProperty += alias + ",\t" + str(objId) + ",\t" + tag + ",\t" + STDType + "\r\n"
```

---

# Spire.Doc Python Content Control
## Lock content control content in Word document
```python
#Create StructureDocumentTag for document
sdt = StructureDocumentTag(doc)
section2 = doc.AddSection()
section2.Body.ChildObjects.Add(sdt)

#Specify the type
sdt.SDTProperties.SDTType = SdtType.RichText
for k in range(section.Body.ChildObjects.Count):
    obj = section.Body.ChildObjects.get_Item(k)
    if obj.DocumentObjectType == DocumentObjectType.Table:
        sdt.SDTContent.ChildObjects.Add(obj.Clone())

#Lock content
sdt.SDTProperties.LockSettings = LockSettingsType.ContentLocked
doc.Sections.Remove(section)
```

---

# Spire.Doc Python SDT Color Modification
## Modify the color of Structured Document Tags in a Word document
```python
# Color setting failed

# Iterate through the sections in the document
for s in range(doc.Sections.Count):
    # Get the current section
    section = doc.Sections[s]

    # Iterate through the child objects in the section's body
    for i in range(section.Body.ChildObjects.Count):
        # Check if the child object is a Paragraph
        if isinstance(section.Body.ChildObjects[i], Paragraph):
            # Get the paragraph object
            para = section.Body.ChildObjects[i] if isinstance(
                section.Body.ChildObjects[i], Paragraph) else None

            # Iterate through the child objects in the paragraph
            for j in range(para.ChildObjects.Count):
                # Check if the child object is a StructureDocumentTagInline
                if isinstance(para.ChildObjects[j],
                              StructureDocumentTagInline):
                    # Get the StructureDocumentTagInline object
                    sdt = para.ChildObjects[j] if isinstance(
                        para.ChildObjects[j],
                        StructureDocumentTagInline) else None

                    # Get the SDTProperties of the StructureDocumentTagInline
                    sDTProperties = sdt.SDTProperties

                    # Set the color of the SDTProperties based on the SDTType
                    if sDTProperties.SDTType == SdtType.RichText:
                        sDTProperties.Color = Color.get_Orange()
                    elif sDTProperties.SDTType == SdtType.Text:
                        sDTProperties.Color = Color.get_Green()

        # Check if the child object is a StructureDocumentTag
        if isinstance(section.Body.ChildObjects[i], StructureDocumentTag):
            # Get the StructureDocumentTag object
            sdt = section.Body.ChildObjects[i] if isinstance(
                section.Body.ChildObjects[i], StructureDocumentTag) else None

            # Get the SDTProperties of the StructureDocumentTag
            sDTProperties = sdt.SDTProperties

            # Set the color of the SDTProperties based on the SDTType
            if sDTProperties.SDTType == SdtType.RichText:
                sDTProperties.Color = Color.Orange
            elif sDTProperties.SDTType == SdtType.Text:
                sDTProperties.Color = Color.Green
```

---

# spire.doc python content controls
## remove structured document tags from word document
```python
#Loop through sections
for s in range(doc.Sections.Count):
    section = doc.Sections[s]
    i = 0
    while i < section.Body.ChildObjects.Count:
        #Loop through contents in paragraph
        if isinstance(section.Body.ChildObjects[i], Paragraph):
            para = section.Body.ChildObjects[i] if isinstance(section.Body.ChildObjects[i], Paragraph) else None
            j = 0
            while j < para.ChildObjects.Count:
                #Find the StructureDocumentTagInline
                if isinstance(para.ChildObjects[j], StructureDocumentTagInline):
                    sdt = para.ChildObjects[j] if isinstance(para.ChildObjects[j], StructureDocumentTagInline) else None
                    #Remove the content control from paragraph
                    para.ChildObjects.Remove(sdt)
                    j -= 1
                j += 1
        if isinstance(section.Body.ChildObjects[i], StructureDocumentTag):
            sdt = section.Body.ChildObjects[i] if isinstance(section.Body.ChildObjects[i], StructureDocumentTag) else None
            section.Body.ChildObjects.Remove(sdt)
            i -= 1
        i += 1
```

---

# spire.doc python structured document tag
## set content control appearance
```python
# Iterate through the sections in the document
for i in range(doc.Sections.Count):
    section = doc.Sections.get_Item(i)
    # Iterate through the child objects in the section's body
    for j in range(section.Body.ChildObjects.Count):
        docObj = section.Body.ChildObjects.get_Item(j)
        # Check if the current object is a StructureDocumentTag
        if isinstance(docObj, StructureDocumentTag):
            # Get the StructureDocumentTag object and its SDTProperties
            stdTag = docObj
            sDTProperties = stdTag.SDTProperties

            # Set the appearance of the StructureDocumentTag based on its SDTType
            if sDTProperties.SDTType == SdtType.Text:
                sDTProperties.Appearance = SdtAppearance.BoundingBox
            elif sDTProperties.SDTType == SdtType.RichText:
                sDTProperties.Appearance = SdtAppearance.Hidden
            elif sDTProperties.SDTType == SdtType.Picture:
                sDTProperties.Appearance = SdtAppearance.Tags
            elif sDTProperties.SDTType == SdtType.CheckBox:
                sDTProperties.Appearance = SdtAppearance.Default
```

---

# Spire.Doc Python Checkbox Update
## Toggle checkbox status in Word document
```python
class StructureTags:
    def __init__(self):
        self._m_tagInlines = None

    def get_tag_inlines(self):
        if self._m_tagInlines is None:
            self._m_tagInlines = []
        return self._m_tagInlines
    def set_tag_inlines(self, value):
        self._m_tagInlines = value

def _GetAllTags(document):
    # Create StructureTags
    structureTags = StructureTags()

    # Traverse document sections
    for i in range(document.Sections.Count):
        section = document.Sections.get_Item(i)
        for j in range(section.Body.ChildObjects.Count):
            obj = section.Body.ChildObjects.get_Item(j)
            # Traverse document paragraphs
            if obj.DocumentObjectType == DocumentObjectType.Paragraph:
                tempParagraph = (obj if isinstance(obj, Paragraph) else None)
                for k in range(tempParagraph.ChildObjects.Count):
                    pobj = tempParagraph.ChildObjects.get_Item(k)
                    # Get StructureDocumentTagInline
                    if pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline:
                        structureTags.get_tag_inlines().append(pobj if isinstance(pobj, StructureDocumentTagInline) else None)

    return structureTags

# Get all structured document tags from the document
structureTags = _GetAllTags(document)

# Create list of tags
tagInlines = structureTags.get_tag_inlines()

# Get the controls and update checkbox status
for item in tagInlines:
    # Get the type
    sdtType = item.SDTProperties.SDTType.name

    # Update the status
    if sdtType == "CheckBox":
        tempPro = item.SDTProperties.ControlProperties
        scb = tempPro if isinstance(tempPro, SdtCheckBox) else None
        if scb.Checked:
            scb.Checked = False
        else:
            scb.Checked = True
```

---

# Spire.Doc Python Endnote
## Insert and format endnote in document
```python
# Get section and paragraph
s = doc.Sections[0]
p = s.Paragraphs[1]

# Add endnote
endnote = p.AppendFootnote(FootnoteType.Endnote)

# Append text
text = endnote.TextBody.AddParagraph().AppendText("Reference: Wikipedia")

# Set text format
text.CharacterFormat.FontName = "Impact"
text.CharacterFormat.FontSize = 14
text.CharacterFormat.TextColor = Color.get_DarkOrange()

# Set marker format of endnote
endnote.MarkerCharacterFormat.FontName = "Calibri"
endnote.MarkerCharacterFormat.FontSize = 25
endnote.MarkerCharacterFormat.TextColor = Color.get_DarkBlue()
```

---

# Spire.Doc Python Footnote
## Insert footnote in Word document
```python
# Find the first matched string
selection = document.FindString("Spire.Doc", False, True)
textRange = selection.GetAsOneRange()
paragraph = textRange.OwnerParagraph
index = paragraph.ChildObjects.IndexOf(textRange)

# Insert footnote
footnote = paragraph.AppendFootnote(FootnoteType.Footnote)
paragraph.ChildObjects.Insert(index + 1, footnote)

# Add and format footnote text
textRange = footnote.TextBody.AddParagraph().AppendText("Welcome to evaluate Spire.Doc")
textRange.CharacterFormat.FontName = "Arial Black"
textRange.CharacterFormat.FontSize = 10
textRange.CharacterFormat.TextColor = Color.get_DarkGray()

# Format footnote marker
footnote.MarkerCharacterFormat.FontName = "Calibri"
footnote.MarkerCharacterFormat.FontSize = 12
footnote.MarkerCharacterFormat.Bold = True
footnote.MarkerCharacterFormat.TextColor = Color.get_DarkGreen()
```

---

# Spire.Doc Python Footnote Removal
## Core functionality for removing footnotes from a Word document
```python
# Get the first section of the document
section = document.Sections[0]
# Traverse paragraphs in the section and find the footnote
for y in range(section.Paragraphs.Count):
    para = section.Paragraphs.get_Item(y)
    index = -1
    i = 0
    cnt = para.ChildObjects.Count
    while i < cnt:
        pBase = para.ChildObjects[i] if isinstance(para.ChildObjects[i], ParagraphBase) else None
        if isinstance(pBase, Footnote):
            index = i
            break
        i += 1
    if index > -1:
        # Remove the footnote
        para.ChildObjects.RemoveAt(index)
```

---

# Spire.Doc Python Footnote Formatting
## Set footnote position and number format
```python
#Get the first section
sec = doc.Sections[0]

#Set the number format, restart rule and position for the footnote
sec.FootnoteOptions.NumberFormat = FootnoteNumberFormat.UpperCaseLetter
sec.FootnoteOptions.RestartRule = FootnoteRestartRule.RestartPage
sec.FootnoteOptions.Position = FootnotePosition.PrintAsEndOfSection
```

---

# Spire.Doc Python VBA Macros
## Detect and remove VBA macros from Word documents
```python
# Create Word document.
document = Document()

# If the document contains Macros, remove them from the document.
if document.IsContainMacro:
    document.ClearMacros()
```

---

# Spire.Doc Python Macros
## Handle documents with macros
```python
# Create a document object
document = Document()
# Load document with macros
document.LoadFromFile(inputFile, FileFormat.Docm)
# Save document with macros
document.SaveToFile(outputFile, FileFormat.Docm)
document.Close()
```

---

# Spire.Doc Python Picture Caption
## Add captions to pictures in a Word document
```python
#Create word document
document = Document()

#Create a new section
section = document.AddSection()

#Add the first picture
par1 = section.AddParagraph()
par1.Format.AfterSpacing = 10.0
pic1 = par1.AppendPicture(inputFile1)

pic1.Height = 100.0
pic1.Width = 120.0
#Add caption to the picture
tempFormat = CaptionNumberingFormat.Number
pic1.AddCaption("Figure", tempFormat, CaptionPosition.BelowItem)

#Add the second picture
par2 = section.AddParagraph()
pic2 = par2.AppendPicture(inputFile2)

pic2.Height = 100.0
pic2.Width = 120.0
#Add caption to the picture
pic2.AddCaption("Figure", tempFormat, CaptionPosition.BelowItem)

#Update fields
document.IsUpdateFields = True
```

---

# spire.doc python table caption
## add caption to table in word document
```python
#Get the first table
body = document.Sections[0].Body
table = body.Tables[0] if isinstance(body.Tables[0], Table) else None

#Add caption to the table
table.AddCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.BelowItem)

#Update fields
document.IsUpdateFields = True
```

---

# Spire.Doc Picture Caption and Cross-Reference
## Create picture captions and cross-references in a Word document
```python
#Create word document
document = Document()

#Create a new section
section = document.AddSection()

#Add the first paragraph
firstPara = section.AddParagraph()

#Add the first picture
par1 = section.AddParagraph()
par1.Format.AfterSpacing = 10.0
# Picture would be loaded here
pic1.Height = 120.0
pic1.Width = 120.0
#Add caption to the picture
tempFormat = CaptionNumberingFormat.Number
captionParagraph1 = pic1.AddCaption("Figure", tempFormat, CaptionPosition.BelowItem)
section.AddParagraph()

#Add the second picture
par2 = section.AddParagraph()
# Picture would be loaded here
pic2.Height = 120.0
pic2.Width = 120.0
#Add caption to the picture
captionParagraph2 = pic2.AddCaption("Figure", tempFormat, CaptionPosition.BelowItem)
section.AddParagraph()

#Create a bookmark
bookmarkName = "Figure_2"
paragraph = section.AddParagraph()
paragraph.AppendBookmarkStart(bookmarkName)
paragraph.AppendBookmarkEnd(bookmarkName)

#Replace bookmark content
navigator = BookmarksNavigator(document)
navigator.MoveToBookmark(bookmarkName)
part = navigator.GetBookmarkContent()
part.BodyItems.Clear()
part.BodyItems.Add(captionParagraph2)
navigator.ReplaceBookmarkContent(part)

#Create cross-reference field to point to bookmark "Figure_2"
field = Field(document)
field.Type = FieldType.FieldRef
field.Code = """REF Figure_2 \p \h"""
firstPara.ChildObjects.Add(field)
fieldSeparator = FieldMark(document, FieldMarkType.FieldSeparator)
firstPara.ChildObjects.Add(fieldSeparator)

#Set the display text of the field
tr = TextRange(document)
tr.Text = "Figure 2"
firstPara.ChildObjects.Add(tr)

fieldEnd = FieldMark(document, FieldMarkType.FieldEnd)
firstPara.ChildObjects.Add(fieldEnd)

#Update fields
document.IsUpdateFields = True
```

---

# Spire.Doc Python Caption with Chapter Number
## Set caption with chapter number for images in Word document
```python
#Get the first section
section = document.Sections[0]

#Label name
name = "Caption "
for i in range(section.Body.Paragraphs.Count):
    for j in range(section.Body.Paragraphs[i].ChildObjects.Count):
        if isinstance(section.Body.Paragraphs[i].ChildObjects[j], DocPicture):
            pic1 = section.Body.Paragraphs[i].ChildObjects[j] if isinstance(section.Body.Paragraphs[i].ChildObjects[j], DocPicture) else None
            body = pic1.OwnerParagraph.Owner if isinstance(pic1.OwnerParagraph.Owner, Body) else None
            if body is not None:
                imageIndex = body.ChildObjects.IndexOf(pic1.OwnerParagraph)
                #Create a new paragraph
                para = Paragraph(document)
                #Set label
                para.AppendText(name)
                #Add caption
                field1 = para.AppendField("test", FieldType.FieldStyleRef)
                #Chapter number
                field1.Code = " STYLEREF 1 \\s "
                #Chapter delimiter
                para.AppendText(" - ")
                #Add picture sequence number
                field2 = para.AppendField(name, FieldType.FieldSequence)
                field2.CaptionName = name
                field2.NumberFormat = CaptionNumberingFormat.Number
                body.Paragraphs.Insert(imageIndex + 1, para)

#Set update fields
document.IsUpdateFields = True
```

---

# Spire.Doc Python Table Caption and Cross-Reference
## Create a table with caption and cross-reference in a Word document
```python
#Create word document
document = Document()

#Get the first section
section = document.AddSection()

#Create a table
table = section.AddTable(True)
table.ResetCells(2, 3)

#Add caption to the table
captionParagraph = table.AddCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.BelowItem)

#Create a bookmark
bookmarkName = "Table_1"
paragraph = section.AddParagraph()
paragraph.AppendBookmarkStart(bookmarkName)
paragraph.AppendBookmarkEnd(bookmarkName)

#Replace bookmark content
navigator = BookmarksNavigator(document)
navigator.MoveToBookmark(bookmarkName)
part = navigator.GetBookmarkContent()
part.BodyItems.Clear()
part.BodyItems.Add(captionParagraph)
navigator.ReplaceBookmarkContent(part)

#Create cross-reference field to point to bookmark "Table_1"
field = Field(document)
field.Type = FieldType.FieldRef
field.Code = """REF Table_1 \\p \\h"""

#Insert field to paragraph
paragraph = section.AddParagraph()
testRange = paragraph.AppendText("This is a table caption cross-reference, ")
testRange.CharacterFormat.FontSize = 14
paragraph.ChildObjects.Add(field)

#Insert FieldSeparator object
fieldSeparator = FieldMark(document, FieldMarkType.FieldSeparator)
paragraph.ChildObjects.Add(fieldSeparator)

#Set display text of the field
tr = TextRange(document)
tr.Text = "Table 1"
tr.CharacterFormat.FontSize = 14
tr.CharacterFormat.TextColor = Color.get_DeepSkyBlue()
paragraph.ChildObjects.Add(tr)

#Insert FieldEnd object to mark the end of the field
fieldEnd = FieldMark(document, FieldMarkType.FieldEnd)
paragraph.ChildObjects.Add(fieldEnd)

#Update fields
document.IsUpdateFields = True
```

---

# spire.doc python fixed layout
## extract document layout information
```python
# Create a new instance of Document
doc = Document()

# Create a FixedLayoutDocument from the document
layoutDoc = FixedLayoutDocument(doc)

# Get the first line in the first column of the first page
line = layoutDoc.Pages[0].Columns[0].Lines[0]

# Get the paragraph that contains the line
para = line.Paragraph

# Get the text content of the first page
pageText = layoutDoc.Pages[0].Text

# Iterate through each page in the FixedLayoutDocument
for i in range(layoutDoc.Pages.Count):
    page = layoutDoc.Pages[i]
    # Get all the lines on the current page
    lines = page.GetChildEntities(LayoutElementType.Line, True)

# Get the lines of the first paragraph
paragraphLines = layoutDoc.GetLayoutEntitiesOfNode(
    (doc.FirstChild).Body.Paragraphs[0])
for i in range(paragraphLines.Count):
    paragraphLine = paragraphLines.get_Item(i)
```

---

# spire.doc python chart
## append bar chart to document
```python
# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Bar chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a bar chart shape to the paragraph with specified width and height
chartShape = newPara.AppendChart(ChartType.Bar, float(400), float(300))
chart = chartShape.Chart

# Get the title of the chart
title = chart.Title

# Set the text of the chart title
title.Text = "My Chart"

# Show the chart title
title.Show = True

# Overlay the chart title on top of the chart
title.Overlay = True
```

---

# spire.doc python chart
## create bubble chart in document
```python
# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Bubble chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a bubble chart shape to the paragraph with specified width and height
shape = newPara.AppendChart(ChartType.Bubble, float(500), float(300))

# Get the chart object from the shape
chart = shape.Chart

# Clear any existing series in the chart
chart.Series.Clear()

# Add a new series to the chart with data points for X, Y, and bubble size values
series = chart.Series.Add("Test Series", [2.9, 3.5, 1.1, 4.0, 4.0],
                          [1.9, 8.5, 2.1, 6.0, 1.5], [9.0, 4.5, 2.5, 8.0, 5.0])
```

---

# Spire.Doc Column Chart Creation
## Create and append a column chart to a Word document
```python
# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Column chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a column chart shape to the paragraph with specified width and height
shape = newPara.AppendChart(ChartType.Column, float(500), float(300))

# Get the chart object from the shape
chart = shape.Chart

# Clear any existing series in the chart
chart.Series.Clear()

# Add a new series to the chart with data points for X values (categories) and Y values
chart.Series.Add("Test Series",
                 ["Word", "PDF", "Excel", "GoogleDocs", "Office"], [
                     float(1900000),
                     float(850000),
                     float(2100000),
                     float(600000),
                     float(1500000)
                 ])

# Set the number format for the Y-axis labels
chart.AxisY.NumberFormat.FormatCode = "#,##0"
```

---

# Spire.Doc Python Chart
## Append Line Chart to Document
```python
# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Line chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a line chart shape to the paragraph with specified width and height
shape = newPara.AppendChart(ChartType.Line, 500.0, 300.0)

# Get the chart object from the shape
chart = shape.Chart

# Get the title of the chart
title = chart.Title

# Set the text of the chart title
title.Text = "My Chart"

# Clear any existing series in the chart
seriesColl = chart.Series
seriesColl.Clear()

# Define categories (X-axis values)
categories = ["C1", "C2", "C3", "C4", "C5", "C6"]

# Add two series to the chart with specified categories and Y-axis values
seriesColl.Add("AW Series 1", categories, [1.0, 2.0, 2.5, 4.0, 5.0, 6.0])
seriesColl.Add("AW Series 2", categories, [2.0, 3.0, 3.5, 6.0, 6.5, 7.0])
```

---

# Spire.Doc Python Pie Chart Creation
## Core functionality for creating a pie chart in a Word document using Spire.Doc for Python
```python
# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Pie chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a pie chart shape to the paragraph with specified width and height
shape = newPara.AppendChart(ChartType.Pie, 500.0, 300.0)
chart = shape.Chart

# Add a series to the chart with categories (labels) and corresponding data values
series = chart.Series.Add("Test Series", ["Word", "PDF", "Excel"],
                          [2.7, 3.2, 0.8])
```

---

# spire.doc python chart
## create scatter chart in word document
```python
# Create a new instance of Document
document = Document()

# Add a section to the document
section = document.AddSection()

# Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Scatter chart.")

# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a scatter chart shape to the paragraph with specified width and height
shape = newPara.AppendChart(ChartType.Scatter, 450.0, 300.0)
chart = shape.Chart

# Clear any existing series in the chart
chart.Series.Clear()

# Add a new series to the chart with data points for X and Y values
chart.Series.Add("Scatter chart", [1.0, 2.0, 3.0, 4.0, 5.0],
                 [1.0, 20.0, 40.0, 80.0, 160.0])
```

---

# Spire.Doc Python Chart
## Create a Surface3D chart in a Word document
```python
# Add a new paragraph to the section
newPara = section.AddParagraph()

# Append a Surface3D chart shape to the paragraph with specified width and height
shape = newPara.AppendChart(ChartType.Surface3D, 500.0, 300.0)

# Get the chart object from the shape
chart = shape.Chart

# Clear any existing series in the chart
chart.Series.Clear()

# Set the title of the chart
chart.Title.Text = "My chart"

# Add multiple series to the chart with categories (X-axis values) and corresponding data values
chart.Series.Add("Series 1", ["Word", "PDF", "Excel", "GoogleDocs", "Office"],
                 [1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0])

chart.Series.Add("Series 2", ["Word", "PDF", "Excel", "GoogleDocs", "Office"],
                 [900000.0, 50000.0, 1100000.0, 400000.0, 2500000.0])

chart.Series.Add("Series 3", ["Word", "PDF", "Excel", "GoogleDocs", "Office"],
                 [500000.0, 820000.0, 1500000.0, 400000.0, 100000.0])
```

---

