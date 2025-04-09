Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace Encrypt
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\Template.docx")

			' Encrypt the document with the provided password
			document.Encrypt("E-iceblue")

			' Save the encrypted document to the specified file path in DOCX format
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			' Dispose the document object to free up resources
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("Sample.docx")


		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
