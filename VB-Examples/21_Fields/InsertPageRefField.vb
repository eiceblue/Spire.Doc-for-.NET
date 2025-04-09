Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections
Imports System.Text

Namespace InsertPageRefField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the document from a specified file path
			Dim document As New Document("..\..\..\..\..\..\Data\PageRef.docx")

			' Get the last section of the document
			Dim section As Section = document.LastSection

			' Add a paragraph to the section
			Dim par As Paragraph = section.AddParagraph()

			' Append a page reference field with the specified name and type
			Dim field As Field = par.AppendField("pageRef", FieldType.FieldPageRef)

			' Set the code for the field with the specified parameters
			field.Code = "PAGEREF bookmark1 \# ""0"" \* Arabic  \* MERGEFORMAT"

			' Enable the automatic update of fields in the document
			document.IsUpdateFields = True

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
