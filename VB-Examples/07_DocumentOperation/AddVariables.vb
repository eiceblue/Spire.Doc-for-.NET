Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace AddVariables
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Add a new section to the document
			Dim section As Section = document.AddSection()

			' Add a new paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Append a field with the specified text and type to the paragraph
			paragraph.AppendField("A1", FieldType.FieldDocVariable)

			' Add a variable with the specified name and value to the document
			document.Variables.Add("A1", "12")

			' Set the IsUpdateFields property to true to update fields when saving the document
			document.IsUpdateFields = True

			' Specify the file path for the output result
			Dim result As String = "Result-AddVariables.docx"

			' Save the document to a file with the specified file format
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document object
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
