from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/ConvertedTemplate.docx"
outputFile = "ToPdfWithPassword.pdf"

#create word document
document = Document()
document.LoadFromFile(inputFile)

#create a parameter
toPdf = ToPdfParameterList()

#set the password
password = "E-iceblue"

toPdf.PdfSecurity.Encrypt("password", password, PdfPermissionsFlags.Default, PdfEncryptionKeySize.Key128Bit)

#save doc file.
document.SaveToFile(outputFile, toPdf)
document.Close()
