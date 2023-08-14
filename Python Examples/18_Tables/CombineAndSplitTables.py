from spire.doc import *
from spire.doc.common import *

##endregion
def SplitTable():
    outputFile = "./Data/SplitTable.docx"
    inputFile =  "./Data/CombineAndSplitTables.docx"

    #Load document from disk
    doc = Document()
    doc.LoadFromFile(inputFile)

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

    #Save the Word file
    section.Document.SaveToFile(outputFile, FileFormat.Docx2013)
    doc.Close()

##endregion
def CombineTables():
    inputFile = "./Data/CombineAndSplitTables.docx"
    outputFile = "CombineTables.docx"

    #Load document from disk
    doc = Document()
    doc.LoadFromFile(inputFile)

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

    #Save the Word file
    section.Document.SaveToFile(outputFile, FileFormat.Docx2013)
    doc.Close()



# Combine tables
CombineTables()

# Split a table
SplitTable()





