Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports System.IO

Namespace CreateTableFromHTML
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Define an HTML string representing a table structure
			Dim HTML As String = "<table border='2px'>" & "<tr>" & "<td>Row 1, Cell 1</td>" & "<td>Row 1, Cell 2</td>" & "</tr>" & "<tr>" & "<td>Row 2, Cell 1</td>" & "<td>Row 2, Cell 2</td>" & "</tr>" & "</table>"

			' Create a new Document object
			Dim document As New Document()

			' Add a new section to the document
			Dim section As Section = document.AddSection()

			' Append the HTML content to the section as a paragraph
			section.AddParagraph().AppendHTML(HTML)

			' Save the document to a file in Docx2013 format
			Dim output As String = "CreateTableFromHTML_out.docx"
			document.SaveToFile(output, FileFormat.Docx2013)

			' Dispose of the document object to free up resources
			document.Dispose()

			'Launch the file
			WordDocViewer(output)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
