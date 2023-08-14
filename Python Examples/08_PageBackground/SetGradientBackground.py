from spire.doc import *
from spire.doc.common import *

inputFile = "./Data/Template_Docx_2.docx"
outputFile = "SetGradientBackground.docx"

#Create Word document.
document = Document()
#Load the file from disk.
document.LoadFromFile(inputFile)
#Set the background type as Gradient.
document.Background.Type = BackgroundType.Gradient
Test = document.Background.Gradient
#Set the first color and second color for Gradient.
Test.Color1 = Color.get_White()
Test.Color2 = Color.get_LightBlue()
#Set the Shading style and Variant for the gradient.
Test.ShadingVariant = GradientShadingVariant.ShadingDown
Test.ShadingStyle = GradientShadingStyle.Horizontal
#Save to file.
document.SaveToFile(outputFile, FileFormat.Docx2013)
document.Close()