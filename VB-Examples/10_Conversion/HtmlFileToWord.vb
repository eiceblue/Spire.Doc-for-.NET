Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace HtmlFileToWord
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the HTML file into the Document object, specifying the input file path, file format as HTML, and XHTML validation type as None
			document.LoadFromFile("..\..\..\..\..\..\..\Data\InputHtmlFile.html", FileFormat.Html, XHTMLValidationType.None)

			' Save the document to a Word file in Docx format with the specified output file name
			document.SaveToFile("HtmlFileToWord.docx", FileFormat.Docx)

			' Dispose of the Document object to release resources
			document.Dispose()


			'Launch the file.
			WordDocViewer("HtmlFileToWord.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
