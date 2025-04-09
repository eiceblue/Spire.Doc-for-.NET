Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections
Imports System.Text

Namespace FormFieldsProperties
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the document from a file
			Dim document As New Document("..\..\..\..\..\..\Data\FillFormField.doc")

			' Get the first section of the document
			Dim section As Section = document.Sections(0)

			' Get the second form field in the section
			Dim formField As FormField = section.Body.FormFields(1)

			' Check if the form field is a text input field
			If formField.Type = FieldType.FieldFormTextInput Then
				' Set the text of the form field
				formField.Text = "My name is " & formField.Name

				' Customize the text formatting of the form field
				formField.CharacterFormat.TextColor = Color.Red
				formField.CharacterFormat.Italic = True
			End If

			' Save the modified document to a file
			document.SaveToFile("result.docx", FileFormat.Docx)

			' Dispose the document object
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
