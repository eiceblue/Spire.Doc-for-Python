#from io import FileIO
from spire.doc import *
from spire.doc.common import *


inputFile = "./Data/TableSample.docx"
outputFile = "GetTablePosition.txt"

# Create a document
document = Document()
# Load file
document.LoadFromFile(inputFile)
# Get the first section
section = document.Sections[0]
# Get the first table
table = section.Tables[0] if isinstance(section.Tables[0], Table) else None

stringBuidler = ''

# Verify whether the table uses "Around" text wrapping or not.
if table.TableFormat.WrapTextAround:
    positon = table.TableFormat.Positioning

    stringBuidler += "Horizontal:"
    stringBuidler += "\n"
    stringBuidler += "Position:" + str(positon.HorizPosition) + " pt"
    stringBuidler += "\n"
    stringBuidler += "Absolute Position:" + positon.HorizPositionAbs + \
        ", Relative to:" + positon.HorizRelationTo
    stringBuidler += "\n"
    stringBuidler += "\n"
    stringBuidler += "Vertical:"
    stringBuidler += "\n"
    stringBuidler += "Position:" + str(positon.VertPosition) + " pt"
    stringBuidler += "\n"
    stringBuidler += "Absolute Position:" + positon.VertPositionAbs + \
        ", Relative to:" + positon.VertRelationTo
    stringBuidler += "\n"
    stringBuidler += "\n"
    stringBuidler += "Distance from surrounding text:"
    stringBuidler += "\n"
    stringBuidler += "Top:" + \
        str(positon.DistanceFromTop) + " pt, Left:" + \
        str(positon.DistanceFromLeft) + " pt"
    stringBuidler += "\n"
    stringBuidler += "Bottom:" + \
        str(positon.DistanceFromBottom) + "pt, Right:" + \
        str(positon.DistanceFromRight) + " pt"
    stringBuidler += "\n"

# Save file.
#FileIO.WriteAllText(outputFile, stringBuidler)
with open(outputFile, "w") as file:
    file.write(stringBuidler)
document.Close()
