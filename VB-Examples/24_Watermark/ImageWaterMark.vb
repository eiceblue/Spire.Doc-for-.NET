Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ImageWaterMark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the document from a template file
			Dim document As New Document("..\..\..\..\..\..\Data\Template.docx")

			' Insert image watermark
			InsertImageWatermark(document)

			' Save the modified document to a new file
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			' Dispose the document object
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("Sample.docx")


		End Sub

		Private Shared Sub InsertImageWatermark(ByVal document As Document)
			' Create a PictureWatermark object
			Dim picture As New PictureWatermark()
			' Load the image for the watermark
			picture.Picture = Image.FromFile("..\..\..\..\..\..\Data\ImageWatermark.png")
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			' picture.SetPicture("..\..\..\..\..\..\Data\ImageWatermark.png")
			' =============================================================================

			' Set the scaling of the watermark
			picture.Scaling = 250
			' Specify whether the watermark should be washed out
			picture.IsWashout = False
			' Set the watermark for the document
			document.Watermark = picture
		End Sub
		
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
