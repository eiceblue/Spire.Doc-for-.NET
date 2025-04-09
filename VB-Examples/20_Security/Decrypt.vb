Imports Spire.Doc

Namespace Decrypt
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
				' Create a new Document object
				Dim document As New Document()

				' Load the document from the specified file path using the provided password
				document.LoadFromFile("..\..\..\..\..\..\Data\TemplateWithPassword.docx", FileFormat.Docx, "E-iceblue")

				' Save the document to the specified file path in DOCX format
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
