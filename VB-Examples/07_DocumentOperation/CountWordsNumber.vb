Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.IO

Namespace CountWordsNumber
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load a document from a file ("..\..\..\..\..\..\Data\Template_Docx_1.docx")
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			' Create a new StringBuilder object to store the content
			Dim content As New StringBuilder()

			' Append the character count to the content
			content.AppendLine("CharCount: " & document.BuiltinDocumentProperties.CharCount)

			' Append the character count with spaces to the content
			content.AppendLine("CharCountWithSpace: " & document.BuiltinDocumentProperties.CharCountWithSpace)

			' Append the word count to the content
			content.AppendLine("WordCount: " & document.BuiltinDocumentProperties.WordCount)

			' Define the file name for the resulting text file
			Dim result As String = "Result-CountWordsNumber.txt"

			' Write the content to a text file
			File.WriteAllText(result, content.ToString())

			' Dispose of the Document object to release resources
			document.Dispose()

			'Launch the file.
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
