Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections

Namespace RemoveField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the document from a specified file path
			Dim document As New Document("..\..\..\..\..\..\Data\IfFieldSample.docx")

			' Get the first field in the document
			Dim field As Field = document.Fields(0)

			' Get the parent paragraph of the field
			Dim par As Paragraph = field.OwnerParagraph

			' Get the index of the field within the child objects of the paragraph
			Dim index As Integer = par.ChildObjects.IndexOf(field)

			' Remove the field from the paragraph
			par.ChildObjects.RemoveAt(index)

			' Save the modified document to a file with the specified name and format
			document.SaveToFile("result.docx", FileFormat.Docx)

			' Dispose of the document object to free up resources
			document.Dispose()

			'Launch the Word file
			WordDocViewer("result.docx")

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
