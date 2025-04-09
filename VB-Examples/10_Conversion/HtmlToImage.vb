Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.Drawing.Imaging

Namespace HtmlToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create Word document.
			Dim document As New Document()

			'Load the file from disk.
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_HtmlFile1.html", FileFormat.Html, XHTMLValidationType.None)

			Dim result As String = "Result-HtmlToImage.png"

			'Save to image
			Dim image As Image = document.SaveToImages(0, ImageType.Bitmap)
			image.Save(result, ImageFormat.Png)

			'Dispose the document
			document.Dispose()

			'Launch the image.
			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
