import locale
from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ComboBox.docx"
outputFile = "ComboBoxItem.docx"

# Create a new document and load from file
doc = Document()
doc.LoadFromFile(inputFile)

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

#Save the document and launch it
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()
