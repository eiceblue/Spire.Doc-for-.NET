Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.Text.RegularExpressions

Namespace RemoveTableOfContent
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load a Word document from a specific file path
			document.LoadFromFile("..\..\..\..\..\..\Data\TableOfContent.docx")

			' Access the body of the first section in the document
			Dim body As Body = document.Sections(0).Body

			' Define a regular expression pattern to match the style names
			Dim regex As New Regex("TOC\w+")

			' Iterate over the paragraphs in the body
			Dim i As Integer = 0
			Do While i < body.Paragraphs.Count
				' Check if the style name matches the regular expression pattern
				If regex.IsMatch(body.Paragraphs(i).StyleName) Then
					' Remove the paragraph if it matches the pattern
					body.Paragraphs.RemoveAt(i)

					' Decrement the counter to avoid skipping the next paragraph
					i -= 1
				End If
				i += 1
			Loop

			' Save the modified document to a new file named "Output.docx"
			document.SaveToFile("Output.docx", FileFormat.Docx)

			' Dispose the document object to free up resources
			document.Dispose()

			'Launch the Word file.
			WordDocViewer("Output.docx")
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
