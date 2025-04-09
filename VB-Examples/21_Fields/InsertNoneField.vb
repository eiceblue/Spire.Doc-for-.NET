Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections
Imports System.Text

Namespace InsertNoneField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the document from a specified file path
			Dim document As New Document("..\..\..\..\..\..\Data\SampleB_2.docx")

			' Get the first section of the document
			Dim section As Section = document.Sections(0)

			' Add a paragraph to the section
			Dim par As Paragraph = section.AddParagraph()

			' Append an empty field with no specific type
			Dim field As Field = par.AppendField(String.Empty, FieldType.FieldNone)

			' Save the modified document to a file with the specified name
			document.SaveToFile("result.docx", FileFormat.Docx)

			' Dispose of the document object to free up resources
			document.Dispose()
			
			'Launch result file
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
