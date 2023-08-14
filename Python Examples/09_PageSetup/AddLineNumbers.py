from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_2.docx"
outputFile = "AddLineNumbers.docx"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Set the start value of the line numbers.
document.Sections[0].PageSetup.LineNumberingStartValue = 1
#Set the interval between displayed numbers.
document.Sections[0].PageSetup.LineNumberingStep = 6
#Set the distance between line numbers and text.
document.Sections[0].PageSetup.LineNumberingDistanceFromText = 40
#Set the numbering mode of line numbers. There are four choices: None, Continuous, RestartPage and RestartSection.
document.Sections[0].PageSetup.LineNumberingRestartMode = LineNumberingRestartMode.Continuous
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()