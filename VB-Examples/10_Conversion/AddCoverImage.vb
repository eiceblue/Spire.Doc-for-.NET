Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Fields
Namespace AddCoverImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim doc As New Document()

			' Load the document from the specified file path
			doc.LoadFromFile("..\..\..\..\..\..\..\Data\ToEpub.doc")

			' Create a new DocPicture object using the document
			Dim picture As New DocPicture(doc)

			' Load the image from the specified file path and assign it to the DocPicture
			picture.LoadImage(Image.FromFile("..\..\..\..\..\..\..\Data\Cover.png"))
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'picture.LoadImage("..\..\..\..\..\..\..\Data\Cover.png")
			' =============================================================================

			' Define the output file name for the EPUB file with the cover image added
			Dim result As String = "AddCoverImage.epub"

			' Save the document to an EPUB file, including the cover image, with the specified output file name
			doc.SaveToEpub(result, picture)

			' Dispose of the Document object to release resources
			doc.Dispose()

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
