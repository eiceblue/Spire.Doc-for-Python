from spire.doc import *

# Create a new empty Word document.
doc = Document()

# Initialize a DocumentNavigator to help insert content and form fields into the document.
navigator = DocumentNavigator(doc)

# Write a label indicating the type of the following form field: Calculation.
navigator.Write("TextFormFieldType.Calculation: ")

# Insert a calculation-type text form field.
navigator.InsertTextFormField("CalculationTextField", TextFormFieldType.Calculation, "0", "=3+1", 30)

# Insert a line break (carriage return) to move to the next line.
navigator.Writeln()

# Write a label for a number-only text form field.
navigator.Write("TextFormFieldType.NumberText: ")

# Insert a number-input text form field.
navigator.InsertTextFormField("NumberText", TextFormFieldType.NumberText, "0", "100", 30)

# Insert a line break.
navigator.Writeln()

# Write a label for a date-input text form field.
navigator.Write("TextFormFieldType.DateText: ")

# Insert a date-formatted text form field.
navigator.InsertTextFormField("DateText", TextFormFieldType.DateText, "yyyy/M/d", "2025/8/1", 30)

# Insert a line break.
navigator.Writeln()

# Enable automatic field updating. 
doc.IsUpdateFields = True

# Save the resulting document to a file.
doc.SaveToFile("InsertTextFormField.docx", FileFormat.Docx)

# Close the document to release internal resources.
doc.Close()

# Explicitly dispose of the document object to free memory immediately.
doc.Dispose()