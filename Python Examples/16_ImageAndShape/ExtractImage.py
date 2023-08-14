import queue
from spire.doc import *
from spire.doc.common import *
import os

outputPath = "ExtractImage/"
inputFile = "./Data/Template.docx"

if not os.path.exists(outputPath):
    os.makedirs(outputPath)

#open document
document = Document()
document.LoadFromFile(inputFile)

#document elements, each of them has child elements
nodes = queue.Queue()
nodes.put(document)

#embedded images list.
images = []

#traverse
while nodes.qsize() > 0:
    node = nodes.get()
    for i in range(node.ChildObjects.Count):
        child = node.ChildObjects.get_Item(i)
        if child.DocumentObjectType == DocumentObjectType.Picture:
            picture = child if isinstance(child, DocPicture) else None
            dataBytes = picture.ImageBytes
            images.append(dataBytes)
        elif isinstance(child, ICompositeObject):
            nodes.put(child if isinstance(child, ICompositeObject) else None)

#Obtain image data in the default format of png,you can use it to convert other image format.
for i, item in enumerate(images):
    fileName = "Image-{}.png".format(i)
    with open(outputPath+fileName,'wb') as imageFile:
        imageFile.write(item)
document.Close()

