from spire.doc import *
from spire.doc.common import *


def _ConvertParagraphToImage(obj):
    doc = Document()
    section = doc.AddSection()

    section.Body.ChildObjects.Add(obj.Clone())
    stream = doc.SaveImageToStreams(0, ImageType.Bitmap)
    doc.Close()
    return stream


def _ConvertTableToImage(obj):
    doc = Document()
    section = doc.AddSection()

    section.Body.ChildObjects.Add(obj.Clone())

    stream = doc.SaveImageToStreams(0, ImageType.Bitmap)
    doc.Close()
    return stream


def _ConvertTableRowToImage(obj):
    doc = Document()
    section = doc.AddSection()
    table = section.AddTable()
    table.Rows.Add(obj.Clone())
    stream = doc.SaveImageToStreams(0, ImageType.Bitmap)
    doc.Close()
    return stream


def _ConvertTableCellToImage(obj):
    doc = Document()
    section = doc.AddSection()
    table = section.AddTable()
    table.AddRow().Cells.Add(obj.Clone())
    stream = doc.SaveImageToStreams(0, ImageType.Bitmap)
    doc.Close()
    return stream


def _ConvertShapeToImage(obj):
    doc = Document()
    section = doc.AddSection()
    section.AddParagraph().ChildObjects.Add(obj.Clone())
    ms = Stream()
    doc.SaveToStream(ms, FileFormat.Docx)
    doc.LoadFromStream(ms, FileFormat.Docx)
    stream = doc.SaveImageToStreams(0, ImageType.Bitmap)
    ms.Close()
    doc.Close()
    return stream


outputFile = "ConvertObjectToImage/"
inputFile = "Data/ConvertObjectToImage.docx"

#Create a document
document = Document()
#Load file
document.LoadFromFile(inputFile)
#Get the first section
section = document.Sections[0]
#Get body of section
body = section.Body

#Get the first paragraph
paragraph = body.Paragraphs[0]
imageStream1 = _ConvertParagraphToImage(paragraph)

imageFile1 = outputFile + "ConvertParagraphToImage.png"
with open(imageFile1, 'wb') as imageFile:
    imageFile.write(imageStream1.ToArray())

#Get the first table
table = body.Tables[0] if isinstance(body.Tables[0], Table) else None

imageStream2 = _ConvertTableToImage(table)
imageFile2 = outputFile + "ConvertTableToImage.jpg"
with open(imageFile2, 'wb') as imageFile:
    imageFile.write(imageStream2.ToArray())

#Get the first row of the first table
row = table.Rows[0]
imageStream3 = _ConvertTableRowToImage(row)

imageFile3 = outputFile + "ConvertTableRowToImage.bmp"
with open(imageFile3, 'wb') as imageFile:
    imageFile.write(imageStream3.ToArray())

#Get the first cell of the first row
cell = row.Cells[0]
imageStream4 = _ConvertTableCellToImage(cell)
#image.Save(outputFile + "ConvertTableCellToImage.png", ImageFormat.Png)
imageFile4 = outputFile + "ConvertTableCellToImage.png"
with open(imageFile4, 'wb') as imageFile:
    imageFile.write(imageStream4.ToArray())

#Get a shape
k = 0
for i in range(0, section.Paragraphs.Count):
    p = section.Paragraphs.get_Item(i)
    for j in range(0, p.ChildObjects.Count):
        obj = p.ChildObjects.get_Item(j)
        if obj.DocumentObjectType == DocumentObjectType.Shape:
            imageStream5 = _ConvertShapeToImage(
                obj if isinstance(obj, ShapeObject) else None)
            #image.Save(string.Format(outputFile + "ConvertShapeToImage-{0}.png", i), ImageFormat.Png)
            imageFile5 = outputFile + "ConvertShapeToImage-" + str(k) + ".png"
            with open(imageFile5, 'wb') as imageFile:
                imageFile.write(imageStream5.ToArray())
            k += 1

#C# TO PYTHON CONVERTER TASK: There is no preprocessor in Python:
##endregion

#def CutImageWhitePart(self, bmp, WhiteBarRate):
#    top = 0
#    left = 0
#    right = bmp.Width
#    bottom = bmp.Height
#    white = Color.White

#    i = 0
#    while i < bmp.Height:
#        find = False
#        j = 0
#        while j < bmp.Width:
#            c = bmp.GetPixel(j, i)
#            if self.IsWhite(c):
#                top = i
#                find = True
#                break
#            j += 1
#        if find:
#            break
#        i += 1

#    i = 0
#    while i < bmp.Width:
#        find = False
#        j = top
#        while j < bmp.Height:
#            c = bmp.GetPixel(i, j)
#            if self.IsWhite(c):
#                left = i
#                find = True
#                break
#            j += 1
#        if find:
#            break
#        pass
#        i += 1

#    for i in range(bmp.Height - 1, -1, -1):
#        find = False
#        j = left
#        while j < bmp.Width:
#            c = bmp.GetPixel(j, i)
#            if self.IsWhite(c):
#                bottom = i
#                find = True
#                break
#            j += 1
#        if find:
#            break

#    for i in range(bmp.Width - 1, -1, -1):
#        find = False
#        for j in range(0, bottom + 1):
#            c = bmp.GetPixel(i, j)
#            if self.IsWhite(c):
#                right = i
#                find = True
#                break
#        if find:
#            break
#    iWidth = right - left
#    iHeight = bottom - left
#    blockWidth = int(math.trunc(iWidth * WhiteBarRate / float(100)))
#    bmp = self.Cut(bmp, left - blockWidth, top - blockWidth, right - left + 2 * blockWidth, bottom - top + 2 * blockWidth)

#    return bmp

#def Cut(self, b, StartX, StartY, iWidth, iHeight):
#    if b is None:
#        return None
#    w = b.Width
#    h = b.Height
#    if StartX >= w or StartY >= h:
#        return None
#    if StartX + iWidth > w:
#        iWidth = w - StartX
#    if StartY + iHeight > h:
#        iHeight = h - StartY
#    try:
#        bmpOut = Bitmap(iWidth, iHeight, PixelFormat.Format24bppRgb)
#        g = Graphics.FromImage(bmpOut)
#        g.DrawImage(b, Rectangle(0, 0, iWidth, iHeight), Rectangle(StartX, StartY, iWidth, iHeight), GraphicsUnit.Pixel)
#        g.Dispose()
#        return bmpOut
#    except:
#        return None
#def IsWhite(self, c):
#    if c.R < 245 or c.G < 245 or c.B < 245:
#        return True
#    else:
#        return False
