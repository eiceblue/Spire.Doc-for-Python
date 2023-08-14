from spire.doc import *
from spire.doc.common import *

outputFile = "CreateTableOfContentByDefault.docx"

doc = Document()
section = doc.AddSection()
para = section.AddParagraph()
#Create table of content with default switches(\o "1-3" \h \z)
para.AppendTOC(1, 3)

par = section.AddParagraph()
tr = par.AppendText("Flowers")
tr.CharacterFormat.FontSize = 30
par.Format.HorizontalAlignment = HorizontalAlignment.Center

#Create paragraph and set the head level
para1 = section.AddParagraph()
para1.AppendText("Ornithogalum")
#Apply the Heading1 style
para1.ApplyStyle(BuiltinStyle.Heading1)
#Add paragraphs
para1 = section.AddParagraph()

picture = para1.AppendPicture("./Data/Ornithogalum.jpg")

picture.TextWrappingStyle = TextWrappingStyle.Square
para1.AppendText("Ornithogalum is a genus of perennial plants mostly native to southern Europe and southern Africa belonging to the family Asparagaceae. Some species are native to other areas such as the Caucasus. Growing from a bulb, species have linear basal leaves and a slender stalk, up to 30 cm tall, bearing clusters of typically white star-shaped flowers, often striped with green.")
para1 = section.AddParagraph()

para2 = section.AddParagraph()
para2.AppendText("Rosa")
#Apply the Heading2 style
para2.ApplyStyle(BuiltinStyle.Heading2)
para2 = section.AddParagraph()
picture2 = para2.AppendPicture("./Data/Rosa.jpg")
picture2.TextWrappingStyle = TextWrappingStyle.Square
para2.AppendText("A rose is a woody perennial flowering plant of the genus Rosa, in the family Rosaceae, or the flower it bears. There are over a hundred species and thousands of cultivars. They form a group of plants that can be erect shrubs, climbing or trailing with stems that are often armed with sharp prickles. Flowers vary in size and shape and are usually large and showy, in colours ranging from white through yellows and reds. Most species are native to Asia, with smaller numbers native to Europe, North America, and northwestern Africa. Species, cultivars and hybrids are all widely grown for their beauty and often are fragrant. Roses have acquired cultural significance in many societies. Rose plants range in size from compact, miniature roses, to climbers that can reach seven meters in height. Different species hybridize easily, and this has been used in the development of the wide range of garden roses.")
section.AddParagraph()
para3 = section.AddParagraph()
para3.AppendText("Hyacinth")
#Apply the Heading3 style
para3.ApplyStyle(BuiltinStyle.Heading3)
para3 = section.AddParagraph()
picture3 = para3.AppendPicture("./Data/hyacinths.JPG")
picture3.TextWrappingStyle = TextWrappingStyle.Tight
para3.AppendText("Hyacinthus is a small genus of bulbous, fragrant flowering plants in the family Asparagaceae, subfamily Scilloideae.These are commonly called hyacinths.The genus is native to the eastern Mediterranean (from the south of Turkey through to northern Israel).")
para3 = section.AddParagraph()
para3.AppendText("Several species of Brodiea, Scilla, and other plants that were formerly classified in the lily family and have flower clusters borne along the stalk also have common names with the word \"hyacinth\" in them. Hyacinths should also not be confused with the genus Muscari, which are commonly known as grape hyacinths.")
#Update TOC
doc.UpdateTableOfContents()
#Save to file
doc.SaveToFile(outputFile, FileFormat.Docx)
doc.Close()

