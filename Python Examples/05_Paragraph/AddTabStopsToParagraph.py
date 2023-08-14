from spire.doc import *
from spire.doc.common import *


outputFile = "AddTabStopsToParagraph.docx"

      
#Create Word document.
document = Document()

#Add a section.
section = document.AddSection()

#Add paragraph 1.
paragraph1 = section.AddParagraph()

#Add tab and set its position (in points).
tab = paragraph1.Format.Tabs.AddTab(28.0)

#Set tab alignment.
tab.Justification = TabJustification.Left

#Move to next tab and append text.
paragraph1.AppendText("\tWashing Machine")

#Add another tab and set its position (in points).
tab = paragraph1.Format.Tabs.AddTab(280.0)

#Set tab alignment.
tab.Justification = TabJustification.Left

#Specify tab leader type.
tab.TabLeader = TabLeader.Dotted

#Move to next tab and append text.
paragraph1.AppendText("\t$650")

#Add paragraph 2.
paragraph2 = section.AddParagraph()

#Add tab and set its position (in points).
tab = paragraph2.Format.Tabs.AddTab(28.0)

#Set tab alignment.
tab.Justification = TabJustification.Left

#Move to next tab and append text.
paragraph2.AppendText("\tRefrigerator")

#Add another tab and set its position (in points).
tab = paragraph2.Format.Tabs.AddTab(280.0)

#Set tab alignment.
tab.Justification = TabJustification.Left

#Specify tab leader type.
tab.TabLeader = TabLeader.NoLeader

#Move to next tab and append text.
paragraph2.AppendText("\t$800")

#Save the Word document
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()