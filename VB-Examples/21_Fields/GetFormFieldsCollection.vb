Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections
Imports System.Text

Namespace GetFormFieldsCollection
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a StringBuilder to hold the field information
			Dim sb As New StringBuilder()

			' Load the document from a file
			Dim document As New Document("..\..\..\..\..\..\Data\FillFormField.doc")

			' Get the first section of the document
			Dim section As Section = document.Sections(0)

			' Get the collection of form fields in the section
			Dim formFields As FormFieldCollection = section.Body.FormFields

			' Append the count of form fields in the section to the StringBuilder
			sb.Append("The first section has " & formFields.Count & " form fields.")

			' Write the result to a text file
			File.WriteAllText("result.txt", sb.ToString())

			' Dispose the document object
			document.Dispose()

			'Launch result file
			WordDocViewer("result.txt")

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
