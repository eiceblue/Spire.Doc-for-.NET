Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AlterLanguageDictionary
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Add a new section to the document
			Dim sec As Section = document.AddSection()

			' Add a new paragraph to the section
			Dim para As Paragraph = sec.AddParagraph()

			' Add text content to the paragraph
			Dim txtRange As TextRange = para.AppendText("corrige seg¨²n diccionario en ingl¨¦s")

			' Set the ASCII locale ID for the character format
			txtRange.CharacterFormat.LocaleIdASCII = 10250

			' Define the file name for the resulting document
			Dim result As String = "Result-AlterLanguageDictionary.docx"

			' Save the document to a file using the specified file format (Docx2013)
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the Document object to release resources
			document.Dispose()

			'Launch the MS Word file.
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
