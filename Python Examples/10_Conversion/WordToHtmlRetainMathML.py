from spire.doc import *

# Initialize a new Document object
document = Document()

# Load an existing Word document from the specified relative file path
document.LoadFromFile("Data\\GetMathEquation.docx")

# Retrieve the HTML export options configuration object for the document
htmlExportOptions = document.HtmlExportOptions

# Configure the export to render Office math equations using MathML format
htmlExportOptions.OfficeMathOutputMode = HtmlOfficeMathOutputMode.MathML

# Set the CSS stylesheet to be embedded internally within the generated HTML file
htmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal

# Define the output file name for the converted HTML document
outputFile = "WordToHtmlRetainMathML.html"

# Save the document as an HTML file using the configured export options
document.SaveToFile(outputFile, FileFormat.Html)

# Close the document to release file handles
document.Close()

# Dispose of the document object to free up memory
document.Dispose()