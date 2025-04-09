Imports Spire.Doc
Imports Spire.Doc.Collections
Imports Spire.Doc.Fields

Namespace ConvertFieldToText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object to store the document
			Dim document As New Document()

			' Load the document from a file
			document.LoadFromFile("..\..\..\..\..\..\Data\Fields.docx")

			' Get the collection of fields in the document
			Dim fields As FieldCollection = document.Fields
			Dim count As Integer = fields.Count

			' Iterate through each field in the collection
			For i As Integer = 0 To count - 1
				' Get the first field in the collection
				Dim field As Field = fields(0)

				' Get the text of the field
				Dim s As String = field.FieldText

				' Get the index of the field within its owner paragraph
				Dim index As Integer = field.OwnerParagraph.ChildObjects.IndexOf(field)

				' Create a TextRange object with the document and set its text to the field text
				Dim textRange As New TextRange(document)
				textRange.Text = s

				' Set the font size of the text range
				textRange.CharacterFormat.FontSize = 24f

				' Insert the text range at the index of the field within its owner paragraph
				field.OwnerParagraph.ChildObjects.Insert(index, textRange)

				' Remove the field from its owner paragraph
				field.OwnerParagraph.ChildObjects.Remove(field)
			Next i

			' Save the modified document to a new file
			document.SaveToFile("ConvertFieldToText.docx", FileFormat.Docx)

			' Dispose the document object
			document.Dispose()

			'Launching the Word file.
			WordDocViewer("ConvertFieldToText.docx")


		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
