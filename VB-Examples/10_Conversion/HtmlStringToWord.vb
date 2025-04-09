Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports System.IO

Namespace HtmlStringToWord
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Read the contents of the HTML file into a string variable
			Dim HTML As String = File.ReadAllText("..\..\..\..\..\..\..\Data\InputHtml.txt")

			' Create a new Document object
			Dim document As New Document()

			' Add a new section to the document
			Dim sec As Section = document.AddSection()

			' Add a paragraph to the section and append the HTML content to it
			sec.AddParagraph().AppendHTML(HTML)

			' Save the document to a Word file in Docx format with the specified output file name
			document.SaveToFile("HtmlFileToWord.docx", FileFormat.Docx)

			' Dispose of the Document object to release resources
			document.Dispose()

			'Launch the Word file.
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
