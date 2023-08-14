from spire.doc import *
from spire.doc.common import *

outputFile = "SetPositionAndNumberFormat.docx"
inputFile = "./Data/Footnote.docx"

#Load the document
doc = Document()
doc.LoadFromFile(inputFile)

#Get the first section
sec = doc.Sections[0]

#Set the number format, restart rule and position for the footnote
sec.FootnoteOptions.NumberFormat = FootnoteNumberFormat.UpperCaseLetter
sec.FootnoteOptions.RestartRule = FootnoteRestartRule.RestartPage
sec.FootnoteOptions.Position = FootnotePosition.PrintAsEndOfSection

#Save and launch document
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()