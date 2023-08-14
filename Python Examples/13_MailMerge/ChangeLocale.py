from spire.doc import *
from spire.doc.common import *
import locale
import datetime

inputFile = "./Data/MailMerge.doc"
outputFile = "ChangeLocale.doc"
#Load word document
document = Document()
document.LoadFromFile(inputFile)

# Store the current culture so it can be set back once mail merge is complete.
current_locale = locale.getlocale()

locale.setlocale(locale.LC_ALL,'de_DE.UTF-8')

fieldNames = ["Contact Name", "Fax", "Date"]
fieldValues = ["John Smith", "+1 (69) 123456", datetime.datetime.now().strftime('%c')]
document.MailMerge.Execute(fieldNames, fieldValues)

locale.setlocale(locale.LC_ALL,current_locale)

document.SaveToFile(outputFile, FileFormat.Doc)
document.Close()
