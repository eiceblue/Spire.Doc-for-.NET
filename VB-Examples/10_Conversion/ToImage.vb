Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ConvertToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
			Dim document As New Document()

			'Load the file from disk
			document.LoadFromFile("..\..\..\..\..\..\Data\ConvertedTemplate.docx")

			'Save the first page to image
			Dim img As Image = document.SaveToImages(0, ImageType.Bitmap)

			'Save to file.
			img.Save("sample.png", Imaging.ImageFormat.Png)

			'Dispose the document
			document.Dispose()

			'Launching the image file.
			WordDocViewer("sample.png")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
