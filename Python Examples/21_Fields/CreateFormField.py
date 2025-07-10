from spire.doc import *
from spire.doc.common import *
import xml.etree.ElementTree as ET

def _InsertHeaderAndFooter(section):
    #insert picture and text to header
    headerParagraph = section.HeadersFooters.Header.AddParagraph()
    headerPicture = headerParagraph.AppendPicture("./Data/Header.png")

    #header text
    text = headerParagraph.AppendText("Demo of Spire.Doc")
    text.CharacterFormat.FontName = "Arial"
    text.CharacterFormat.FontSize = 10
    text.CharacterFormat.Italic = True
    headerParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

    #border
    headerParagraph.Format.Borders.Bottom.BorderType = BorderStyle.Single
    headerParagraph.Format.Borders.Bottom.Space = 0.05


    #header picture layout - text wrapping
    headerPicture.TextWrappingStyle = TextWrappingStyle.Behind

    #header picture layout - position
    headerPicture.HorizontalOrigin = HorizontalOrigin.Page
    headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
    headerPicture.VerticalOrigin = VerticalOrigin.Page
    headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top

    #insert picture to footer
    footerParagraph = section.HeadersFooters.Footer.AddParagraph()
    footerPicture = footerParagraph.AppendPicture("./Data/Footer.png")

    #footer picture layout
    footerPicture.TextWrappingStyle = TextWrappingStyle.Behind
    footerPicture.HorizontalOrigin = HorizontalOrigin.Page
    footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
    footerPicture.VerticalOrigin = VerticalOrigin.Page
    footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom

    #insert page number
    footerParagraph.AppendField("page number", FieldType.FieldPage)
    footerParagraph.AppendText(" of ")
    footerParagraph.AppendField("number of pages", FieldType.FieldNumPages)
    footerParagraph.Format.HorizontalAlignment = HorizontalAlignment.Right

    #border
    footerParagraph.Format.Borders.Top.BorderType = BorderStyle.Single
    footerParagraph.Format.Borders.Top.Space = 0.05

def _AddTitle(section):
  title = section.AddParagraph()
  titleText = title.AppendText("Create Your Account")
  titleText.CharacterFormat.FontSize = 18
  titleText.CharacterFormat.FontName = "Arial"
  titleText.CharacterFormat.TextColor = Color.FromArgb(0xFF, 0x00, 0x71, 0xb6)
  title.Format.HorizontalAlignment = HorizontalAlignment.Center
  title.Format.AfterSpacing = 8

def _SetPage(section):
    #the unit of all measures below is point, 1point = 0.3528 mm
    section.PageSetup.PageSize = PageSize.A4()
    section.PageSetup.Margins.Top = 72
    section.PageSetup.Margins.Bottom = 72
    section.PageSetup.Margins.Left = 89.85
    section.PageSetup.Margins.Right = 89.85

def _AddForm(section):
    descriptionStyle = ParagraphStyle(section.Document)
    descriptionStyle.Name = "description"
    descriptionStyle.CharacterFormat.FontSize = 12
    descriptionStyle.CharacterFormat.FontName = "Arial"
    descriptionStyle.CharacterFormat.TextColor = Color.FromArgb(0xFF, 0x00, 0x45, 0x8e)
    section.Document.Styles.Add(descriptionStyle)

    p1 = section.AddParagraph()
    text1 = "So that we can verify your identity and find your information, " + "please provide us with the following information. " + "This information will be used to create your online account. " + "Your information is not public, shared in anyway, or displayed on this site"
    p1.AppendText(text1)
    p1.ApplyStyle(descriptionStyle.Name)

    p2 = section.AddParagraph()
    text2 = "You must provide a real email address to which we will send your password."
    p2.AppendText(text2)
    p2.ApplyStyle(descriptionStyle.Name)
    p2.Format.AfterSpacing = 8

    #field group label style
    formFieldGroupLabelStyle = ParagraphStyle(section.Document)
    formFieldGroupLabelStyle.Name = "formFieldGroupLabel"
    formFieldGroupLabelStyle.ApplyBaseStyle("description")
    formFieldGroupLabelStyle.CharacterFormat.Bold = True
    formFieldGroupLabelStyle.CharacterFormat.TextColor = Color.get_White()
    section.Document.Styles.Add(formFieldGroupLabelStyle)

    #field label style
    formFieldLabelStyle = ParagraphStyle(section.Document)
    formFieldLabelStyle.Name = "formFieldLabel"
    formFieldLabelStyle.ApplyBaseStyle("description")
    formFieldLabelStyle.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right
    section.Document.Styles.Add(formFieldLabelStyle)

    #add table
    table = section.AddTable()

    #2 columns of per row
    table.DefaultColumnsNumber = 2

    #default height of row is 20point
    table.DefaultRowHeight = 20

    #load form config data
    tree = ET.parse("./Data/Form.xml")
    root = tree.getroot()

    sectionNodes = root.findall(".//section")

    for node in sectionNodes:
        #create a row for field group label, does not copy format
        row = table.AddRow(False)
        row.Cells[0].CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(0xFF, 0x00, 0x71, 0xb6)
        row.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle

        #label of field group
        cellParagraph = row.Cells[0].AddParagraph()
        tempNameStr = node.get("name", "")
        cellParagraph.AppendText(tempNameStr)

        cellParagraph.ApplyStyle(formFieldGroupLabelStyle.Name)

        fieldNodes = node.findall(".//field")
        for fieldNode in fieldNodes:
            #create a row for field, does not copy format
            fieldRow = table.AddRow(False)

            #field label
            fieldRow.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle
            labelParagraph = fieldRow.Cells[0].AddParagraph()
            labelParagraph.AppendText(fieldNode.get("label", ""))
            labelParagraph.ApplyStyle(formFieldLabelStyle.Name)

            fieldRow.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle
            fieldParagraph = fieldRow.Cells[1].AddParagraph()
            fieldId = fieldNode.get("id", "")
            if fieldNode.get("type", "") == "text":
                #add text input field
                fieldFormTextInput = fieldParagraph.AppendField(fieldId, FieldType.FieldFormTextInput)
                field = fieldFormTextInput if isinstance(fieldFormTextInput, TextFormField) else None

                #set default text
                field.DefaultText = ""
                field.Text = ""

            elif fieldNode.get("type", "") == "list":
                #add dropdown field
                fieldFormDropDown = fieldParagraph.AppendField(fieldId, FieldType.FieldFormDropDown)
                fieldList = fieldFormDropDown if isinstance(fieldFormDropDown, DropDownFormField) else None

                #add items into dropdown.
                itemNodes = fieldNode.findall(".//item")
                for itemNode in itemNodes:
                    fieldList.DropDownItems.Add(itemNode.text)

            elif fieldNode.get("type", "") == "checkbox":
                #add checkbox field
                fieldParagraph.AppendField(fieldId, FieldType.FieldFormCheckBox)

        #merge field group row. 2 columns to 1 column
    table.ApplyHorizontalMerge(row.GetRowIndex(), 0, 1)


outputFile="CreateFormField.docx"

#CreateWorddocument.
document=Document()
section=document.AddSection()

#Pagesetup.
_SetPage(section)

#Insertheaderandfooter.
_InsertHeaderAndFooter(section)

#Addtitle.
_AddTitle(section)

#Addform.
_AddForm(section)

#Savedocfile.
document.SaveToFile(outputFile,FileFormat.Docx)
document.Close()