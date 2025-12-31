Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SetImageBackground
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class with the specified file path
			Dim document As New Document("..\..\..\..\..\..\Data\Template.docx")

			' Set the background type of the document to picture
			document.Background.Type = BackgroundType.Picture

			' Set the picture for the document background from the specified image file
			document.Background.Picture = Image.FromFile("..\..\..\..\..\..\Data\Background.png")
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'document.Background.Picture("..\..\..\..\..\..\Data\Background.png")
			' =============================================================================

			' Save the document to a file with the specified file format
			document.SaveToFile("ImageBackground.docx", FileFormat.Docx)

			' Dispose of the document object
			document.Dispose()

			'launching the Word file.
			WordDocViewer("ImageBackground.docx")


		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
