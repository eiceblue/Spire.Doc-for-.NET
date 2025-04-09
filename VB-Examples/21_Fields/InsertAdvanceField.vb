Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections
Imports System.Text

Namespace InsertAdvanceField
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

			' Append a field with the specified type and text
			Dim field As Field = par.AppendField("Field", FieldType.FieldAdvance)

			' Set the code for the field using the specified parameters
			field.Code = "ADVANCE \d 10 \l 10 \r 10 \u 0 \x 100 \y 100 "

			' Enable the automatic update of fields in the document
			document.IsUpdateFields = True

			' Save the modified document to a file with the specified name
			Dim result As String = "result.docx"
			document.SaveToFile(result, FileFormat.Docx)

			' Dispose of the document object to free up resources
			document.Dispose() 
			
			'Launch result file
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
