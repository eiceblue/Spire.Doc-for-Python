from spire.doc import *
from spire.doc.common import *

outputFile = "LockSpecifiedSections.docx"

# Create Word document.
document = Document()

# Add new sections.
s1 = document.AddSection()
s2 = document.AddSection()

# Append some text to section 1 and section 2.
s1.AddParagraph().AppendText("Spire.Doc demo, section 1")
s2.AddParagraph().AppendText("Spire.Doc demo, section 2")

# Protect the document with AllowOnlyFormFields protection type.
document.Protect(ProtectionType.AllowOnlyFormFields, "123")

# Unprotect section 2
s2.ProtectForm = False

# Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()
