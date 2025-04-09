Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections

Namespace ConvertIfFieldToText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object and load the document from a file
			Dim document As New Document("..\..\..\..\..\..\Data\IfFieldSample.docx")

			' Get the collection of fields in the document
			Dim fields As FieldCollection = document.Fields

			' Iterate through each field in the collection
			For i As Integer = 0 To fields.Count - 1
				' Get the current field
				Dim field As Field = fields(i)

				' Check if the field is of type FieldIf
				If field.Type = FieldType.FieldIf Then
					' Cast the field as TextRange to access its properties
					Dim original As TextRange = TryCast(field, TextRange)

					' Get the text of the field
					Dim text As String = field.FieldText

					' Create a new TextRange object with the document and set its text to the field text
					Dim textRange As New TextRange(document)
					textRange.Text = text

					' Set the font name and size of the new text range to match the original field
					textRange.CharacterFormat.FontName = original.CharacterFormat.FontName
					textRange.CharacterFormat.FontSize = original.CharacterFormat.FontSize

					' Get the owner paragraph of the field
					Dim par As Paragraph = field.OwnerParagraph

					' Get the index of the field within its owner paragraph
					Dim index As Integer = par.ChildObjects.IndexOf(field)

					' Remove the field from its owner paragraph
					par.ChildObjects.RemoveAt(index)

					' Insert the new text range at the index of the field within its owner paragraph
					par.ChildObjects.Insert(index, textRange)
				End If
			Next i

			' Specify the file name for the result document
			Dim result As String = "result.docx"

			' Save the modified document to a new file
			document.SaveToFile(result, FileFormat.Docx)

			' Dispose the document object
			document.Dispose()

			'Launch the Word file
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
