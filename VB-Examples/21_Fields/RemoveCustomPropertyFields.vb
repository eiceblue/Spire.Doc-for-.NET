Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RemoveCustomPropertyFields
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim document As New Document()

			' Load an existing document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\RemoveCustomPropertyFields.docx")

			' Get the collection of custom document properties
			Dim cdp As CustomDocumentProperties = document.CustomDocumentProperties

			' Iterate through the custom document properties and remove them
			Dim i As Integer = 0
			Do While i < cdp.Count
				cdp.Remove(cdp(i).Name)
			Loop

			' Enable the automatic update of fields in the document
			document.IsUpdateFields = True

			' Specify the name for the resulting document file
			Dim result As String = "Result-RemoveCustomPropertyFields.docx"

			' Save the modified document to a file with the specified name and format
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document object to free up resources
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
