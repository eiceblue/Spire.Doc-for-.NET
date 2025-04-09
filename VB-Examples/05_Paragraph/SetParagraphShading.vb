Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetParagraphShading
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load the Word document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			' Retrieve the first paragraph in the first section of the document
			Dim paragaph As Paragraph = document.Sections(0).Paragraphs(0)

			' Set the background color of the paragraph to yellow
			paragaph.Format.BackColor = Color.Yellow

			' Retrieve the third paragraph in the first section of the document
			paragaph = document.Sections(0).Paragraphs(2)

			' Find the text "Christmas" within the paragraph
			Dim selection As TextSelection = paragaph.Find("Christmas", True, False)

			' Get the range representing the found text as a single range
			Dim range As TextRange = selection.GetAsOneRange()

			' Set the text background color of the range to yellow
			range.CharacterFormat.TextBackgroundColor = Color.Yellow

			' Specify the output file name for the modified document
			Dim result As String = "Result-SetParagraphShading.docx"

			' Save the modified document to the specified file format (Docx2013)
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document object to release resources
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
