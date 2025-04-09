Imports System.Text
Imports Spire.Doc
Imports System.Text.RegularExpressions

Namespace ReplaceTextByRegex
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object and load the document from the specified file path
			Dim doc As New Document()
			doc.LoadFromFile("..\..\..\..\..\..\Data\ReplaceTextByRegex.docx")

			' Create a regular expression pattern to match words starting with '#' (\#\w+\b)
			Dim regex As New Regex("\#\w+\b")

			' Replace all occurrences of the matched pattern with "Spire.Doc"
			doc.Replace(regex, "Spire.Doc")

			' Save the modified document to a file named "output.docx" in Docx format
			doc.SaveToFile("output.docx", FileFormat.Docx)

			' Dispose of the Document object to release resources
			doc.Dispose()

			'view the document
			WordDocViewer("output.docx")

		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
