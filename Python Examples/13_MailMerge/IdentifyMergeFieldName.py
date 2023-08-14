from io import FileIO
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/IdentifyMergeFieldNames.docx"
outputFile = "IdentifyMergeFieldName.txt"
      
#Create Word document.
document = Document()

#Load the file from disk.
document.LoadFromFile(inputFile)

#Get the collection of group names.
GroupNames = document.MailMerge.GetMergeGroupNames()

#Get the collection of merge field names in a specific group.
MergeFieldNamesWithinRegion = document.MailMerge.GetMergeFieldNames("Products")

#Get the collection of all the merge field names.
MergeFieldNames = document.MailMerge.GetMergeFieldNames()

content = ''
content += "----------------Group Names-----------------------------------------"
content += '\n'
i = 0
while i < len(GroupNames):
    content += GroupNames[i]
    content += '\n'
    i += 1

content += "----------------Merge field names within a specific group-----------"
content += '\n'
j = 0
while j < len(MergeFieldNamesWithinRegion):
    content += MergeFieldNamesWithinRegion[j]
    content += '\n'
    j += 1

content += "----------------All of the merge field names------------------------"
content += '\n'
k = 0
while k < len(MergeFieldNames):
    content += MergeFieldNames[k]
    content += '\n'
    k += 1

#Write the contents in a TXT file
with FileIO(outputFile, mode="w") as f:
    f.write(content.encode("utf-8"))

