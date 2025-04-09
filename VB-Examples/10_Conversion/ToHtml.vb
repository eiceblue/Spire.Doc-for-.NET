Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ConvertToHtml
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\..\Data\ToHtmlTemplate.docx")

			' Save the document to an HTML file with the specified output file name
			document.SaveToFile("Sample.html", FileFormat.Html)

			' Dispose of the Document object to release resources
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("Sample.html")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
