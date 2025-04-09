Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace InsertSymbol
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Add a new Section to the document
			Dim section As Section = document.AddSection()

			' Add a new Paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Use a Unicode character to create the symbol Ä and append it to the paragraph
			Dim tr As TextRange = paragraph.AppendText(ChrW(&H00c4).ToString())

			' Set the text color of the symbol Ä to red
			tr.CharacterFormat.TextColor = Color.Red

			' Add the symbol Ë to the paragraph
			paragraph.AppendText(ChrW(&H00cb).ToString())

			' Specify the output file path
			Dim result As String = "Result-InsertSymbol.docx"

			' Save the document to the specified output file in DOCX2013 format
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
