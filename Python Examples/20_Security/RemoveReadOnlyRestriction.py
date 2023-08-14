from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/RemoveReadOnlyRestriction.docx"
outputFile = "RemoveReadOnlyRestriction.docx"


doc = Document()
doc.LoadFromFile(inputFile)
# Remove ReadOnly Restriction.
doc.Protect(ProtectionType.NoProtection)
doc.SaveToFile(outputFile, FileFormat.Docx2013)
doc.Close()
