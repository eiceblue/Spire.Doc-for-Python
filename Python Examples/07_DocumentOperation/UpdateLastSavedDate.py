from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Sample.docx"
outputFile = "UpdateLastSavedDate.docx"

def LocalTimeToGreenwishTime(lacalTime):
    localTimeZone = TimeZone.get_CurrentTimeZone()
    timeSpan = localTimeZone.GetUtcOffset(lacalTime)
    greenwishTime = lacalTime - timeSpan
    return greenwishTime

#Create Word document.
document = Document()
#Load the document from disk
document.LoadFromFile(inputFile)
#Update the last saved date
document.BuiltinDocumentProperties.LastSaveDate = LocalTimeToGreenwishTime(DateTime.get_Now())
document.SaveToFile(outputFile, FileFormat.Docx)
document.Close()