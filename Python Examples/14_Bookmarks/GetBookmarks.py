from io import FileIO
from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/Bookmarks.docx"
outputFile = "GetBookmarks.txt"
#Create word document
#Load the document from disk.
document = Document()
document.LoadFromFile(inputFile)

#Get the bookmark by index.
bookmark1 = document.Bookmarks[0]

#Get the bookmark by name.
bookmark2 = document.Bookmarks["Test2"]

#Create StringBuilder to save 
content = ''

#Set string format for displaying
result = "The bookmark obtained by index is " + bookmark1.Name + ".\r\nThe bookmark obtained by name is " + bookmark2.Name + ".\n"

#Add result string to StringBuilder
content += result
content += '\n'

#Write the contents in a TXT file
with FileIO(outputFile, mode="w") as f:
    f.write(content.encode("utf-8"))
document.Close()

