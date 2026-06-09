from spire.doc import *

# Create a new instance of the Document class
document = Document()

# Load an existing Word document
document.LoadFromFile("Data\\ConvertedToXLSX.docx")

# Define the file path and name for the output document
result = "ToXLSX.xlsx"

# Convert the Word document to XLSX file
document.SaveToFile(result, FileFormat.XLSX)

# Close the document to release resources
document.Close()

# Dispose of the document object to free up memory
document.Dispose()