Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO
Imports Spire.Doc.Fields
Imports System.Drawing.Imaging

Namespace ConvertObjectToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			
			' Create a new document object
			Dim document As New Document()

			' Load the source document from file
			document.LoadFromFile("..\..\..\..\..\..\Data\ConvertObjectToImage.docx")

			' Get the first section of the document
			Dim section As Section = document.Sections(0)

			' Get the body of the section
			Dim body As Body = section.Body

			' Convert the first paragraph in the body to an image
			Dim paragraph As Paragraph = body.Paragraphs(0)
			Dim image As Image = ConvertParagraphToImage(paragraph)
			image.Save("ConvertParagraphToImage.png", ImageFormat.Png)

			' Convert the first table in the body to an image
			Dim table As Table = TryCast(body.Tables(0), Table)
			image = ConvertTableToImage(table)
			image.Save("ConvertTableToImage.jpg", ImageFormat.Jpeg)

			' Convert the first row in the table to an image
			Dim row As TableRow = table.Rows(0)
			image = ConvertTableRowToImage(row)
			image.Save("ConvertTableRowToImage.bmp", ImageFormat.Bmp)

			' Convert the first cell in the row to an image
			Dim cell As TableCell = row.Cells(0)
			image = ConvertTableCellToImage(cell)
			image.Save("ConvertTableCellToImage.png", ImageFormat.Png)

			' Iterate through paragraphs in the section and convert shape objects to images
			Dim i As Integer = 0
			For Each p As Paragraph In section.Paragraphs
				For Each obj As DocumentObject In p.ChildObjects
					If obj.DocumentObjectType = DocumentObjectType.Shape Then
						image = ConvertShapeToImage(TryCast(obj, ShapeObject))
						image.Save(String.Format("ConvertShapeToImage-{0}.png", i), ImageFormat.Png)
						i += 1
					End If
				Next obj
			Next p

			' Dispose the document
			document.Dispose()
		End Sub
		Private Function ConvertParagraphToImage(ByVal obj As Paragraph) As Image
			' Create a new document and section
			Dim doc As New Document()
			Dim section As Section = doc.AddSection()

			' Add the cloned paragraph to the section's body
			section.Body.ChildObjects.Add(obj.Clone())

			' Save the document to an image and close it
			Dim image As Image = doc.SaveToImages(0, ImageType.Bitmap)
			doc.Close()

			' Cut the white parts from the image
			Return CutImageWhitePart(TryCast(image, Bitmap), 1)
		End Function
		
		Private Function ConvertTableToImage(ByVal obj As Table) As Image
			' Create a new document
			Dim doc As New Document()
			' Add a section to the document
			Dim section As Section = doc.AddSection()

			' Add a clone of the table object to the section's body
			section.Body.ChildObjects.Add(obj.Clone())

			' Save the document as an image and retrieve the image object
			Dim image As Image = doc.SaveToImages(0, ImageType.Bitmap)

			' Close the document
			doc.Close()

			' Return the image after cutting the white parts
			Return CutImageWhitePart(TryCast(image, Bitmap), 1)
		End Function

		Private Function ConvertTableRowToImage(ByVal obj As TableRow) As Image
			' Create a new document
			Dim doc As New Document()
			' Add a section to the document
			Dim section As Section = doc.AddSection()
			' Add a table to the section
			Dim table As Table = section.AddTable()

			' Add a clone of the table row object to the table
			table.Rows.Add(obj.Clone())

			' Save the document as an image and retrieve the image object
			Dim image As Image = doc.SaveToImages(0, ImageType.Bitmap)

			' Close the document
			doc.Close()

			' Return the image after cutting the white parts
			Return CutImageWhitePart(TryCast(image, Bitmap), 1)
		End Function

		Private Function ConvertTableCellToImage(ByVal obj As TableCell) As Image
			' Create a new document
			Dim doc As New Document()
			' Add a section to the document
			Dim section As Section = doc.AddSection()
			' Add a table to the section
			Dim table As Table = section.AddTable()

			' Add a clone of the table cell object to a new row in the table
			table.AddRow().Cells.Add(obj.Clone())

			' Save the document as an image and retrieve the image object
			Dim image As Image = doc.SaveToImages(0, ImageType.Bitmap)

			' Close the document
			doc.Close()

			' Return the image after cutting the white parts
			Return CutImageWhitePart(TryCast(image, Bitmap), 1)
		End Function

		Private Function ConvertShapeToImage(ByVal obj As ShapeObject) As Image
			' Create a new document
			Dim doc As New Document()
			' Add a section to the document
			Dim section As Section = doc.AddSection()
			' Add a paragraph to the section's body
			section.AddParagraph().ChildObjects.Add(obj.Clone())

			' Save the document to a memory stream
			Dim ms As New MemoryStream()
			doc.SaveToStream(ms, FileFormat.Docx)

			' Load the document from the memory stream
			doc.LoadFromStream(ms, FileFormat.Docx)

			' Save the document as an image and retrieve the image object
			Dim image As Image = doc.SaveToImages(0, ImageType.Bitmap)

			' Close the memory stream
			ms.Close()
			' Close the document
			doc.Close()

			' Return the image after cutting the white parts
			Return CutImageWhitePart(TryCast(image, Bitmap), 1)
		End Function

		Public Function CutImageWhitePart(ByVal bmp As Bitmap, ByVal WhiteBarRate As Integer) As Image
			' Initialize variables for the boundaries of the white part
			Dim top As Integer = 0, left As Integer = 0
			Dim right As Integer = bmp.Width, bottom As Integer = bmp.Height
			Dim white As Color = Color.White

			' Find the top boundary of the white part
			For i As Integer = 0 To bmp.Height - 1
				Dim find As Boolean = False
				For j As Integer = 0 To bmp.Width - 1
					Dim c As Color = bmp.GetPixel(j, i)
					If IsWhite(c) Then
						top = i
						find = True
						Exit For
					End If
				Next j
				If find Then
					Exit For
				End If
			Next i

			' Find the left boundary of the white part
			For i As Integer = 0 To bmp.Width - 1
				Dim find As Boolean = False
				For j As Integer = top To bmp.Height - 1
					Dim c As Color = bmp.GetPixel(i, j)
					If IsWhite(c) Then
						left = i
						find = True
						Exit For
					End If
				Next j
				If find Then
					Exit For
				End If
			Next i

			' Find the bottom boundary of the white part
			For i As Integer = bmp.Height - 1 To 0 Step -1
				Dim find As Boolean = False
				For j As Integer = left To bmp.Width - 1
					Dim c As Color = bmp.GetPixel(j, i)
					If IsWhite(c) Then
						bottom = i
						find = True
						Exit For
					End If
				Next j
				If find Then
					Exit For
				End If
			Next i

			' Find the right boundary of the white part
			For i As Integer = bmp.Width - 1 To 0 Step -1
				Dim find As Boolean = False
				For j As Integer = 0 To bottom
					Dim c As Color = bmp.GetPixel(i, j)
					If IsWhite(c) Then
						right = i
						find = True
						Exit For
					End If
				Next j
				If find Then
					Exit For
				End If
			Next i

			' Calculate dimensions for cropping the image
			Dim iWidth As Integer = right - left
			Dim iHeight As Integer = bottom - top
			Dim blockWidth As Integer = Convert.ToInt32(iWidth * WhiteBarRate \ 100)

			' Crop the image based on the calculated boundaries and block width
			bmp = Cut(bmp, left - blockWidth, top - blockWidth, right - left + 2 * blockWidth, bottom - top + 2 * blockWidth)

			Return bmp
		End Function

		' Function to crop an image
		Public Function Cut(ByVal b As Bitmap, ByVal StartX As Integer, ByVal StartY As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer) As Bitmap
			' Check if input bitmap is valid
			If b Is Nothing Then
				Return Nothing
			End If
			Dim w As Integer = b.Width
			Dim h As Integer = b.Height

			' Check if starting coordinates are within bounds
			If StartX >= w OrElse StartY >= h Then
				Return Nothing
			End If

			' Adjust dimensions if exceeding the image size
			If StartX + iWidth > w Then
				iWidth = w - StartX
			End If
			If StartY + iHeight > h Then
				iHeight = h - StartY
			End If

			Try
				' Create a new bitmap with the specified dimensions
				Dim bmpOut As New Bitmap(iWidth, iHeight, PixelFormat.Format24bppRgb)
				Dim g As Graphics = Graphics.FromImage(bmpOut)

				' Draw the cropped portion onto the new bitmap
				g.DrawImage(b, New Rectangle(0, 0, iWidth, iHeight), New Rectangle(StartX, StartY, iWidth, iHeight), GraphicsUnit.Pixel)
				g.Dispose()
				Return bmpOut
			Catch
				Return Nothing
			End Try
		End Function

		' Function to check if a color is considered white
		Public Function IsWhite(ByVal c As Color) As Boolean
			' Check if any RGB component of the color is less than 245 (on a scale of 0-255)
			If c.R < 245 OrElse c.G < 245 OrElse c.B < 245 Then
						Return True
					Else
						Return False
					End If
				End Function
	End Class
End Namespace
