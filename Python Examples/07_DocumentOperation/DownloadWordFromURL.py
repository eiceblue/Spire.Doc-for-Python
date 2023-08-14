from spire.doc import *
from spire.doc.common import *
import requests


outputFile = "DownloadWordFromURL.docx"

#Create Word document.
document = Document()


response = requests.get("https://www.e-iceblue.com/images/test.docx")
if response.status_code == 200:
    content = response.content
    ms = Stream(content)
    document.LoadFromStream(ms, FileFormat.Docx)
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()