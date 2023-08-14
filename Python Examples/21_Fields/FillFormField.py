from spire.doc import *
from spire.doc.common import *
import xml.etree.ElementTree as ET


inputFile = "./Data/FillFormField.doc"
outputFile = "FillFormField.doc"
#Open a Word document with form.
document = Document()
document.LoadFromFile(inputFile)

tree = ET.parse('./Data/User.xml')

root = tree.getroot()


#Fill data.
for k in range(document.Sections[0].Body.FormFields.Count):
    field = document.Sections[0].Body.FormFields.get_Item(k)
    path = "./{0}".format(field.Name.strip())
    propertyNode = root.find(path)
    if propertyNode is not None:
        if field.Type == FieldType.FieldFormTextInput:
            field.Text = propertyNode.text

        elif field.Type == FieldType.FieldFormDropDown:
            combox = field if isinstance(field, DropDownFormField) else None
            for i in range(combox.DropDownItems.Count):
                if combox.DropDownItems[i].Text == propertyNode.text:
                    combox.DropDownSelectedIndex = i
                    break
                if field.Name == "country" and combox.DropDownItems[i].Text == "Others":
                    combox.DropDownSelectedIndex = i

                elif field.Type == FieldType.FieldFormCheckBox:
                    if bool(propertyNode.text):
                        checkBox = field if isinstance(field, CheckBoxFormField) else None
                        checkBox.Checked = True

#Save doc file.
document.SaveToFile(outputFile, FileFormat.Doc)
document.Close()