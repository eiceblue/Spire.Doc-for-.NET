Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections
Imports System.Text

Namespace SetLocaleForField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the document from a file
			Dim document As New Document("..\..\..\..\..\..\Data\SampleB_2.docx")

			' Get the first section of the document
			Dim section As Section = document.Sections(0)

			' Add a paragraph to the section
			Dim par As Paragraph = section.AddParagraph()

			' Append a date field to the paragraph
			Dim field As Field = par.AppendField("DocDate", FieldType.FieldDate)

			' Set the locale ID to Russian (1049) for the first character range in the field
			TryCast(field.OwnerParagraph.ChildObjects(0), TextRange).CharacterFormat.LocaleIdASCII = 1049

			' Set the field text to "2019-10-10"
			field.FieldText = "2019-10-10"

			' Enable automatic update of fields in the document
			document.IsUpdateFields = True

			' Specify the output file name
			Dim result As String = "result.docx"

			' Save the modified document to a new file
			document.SaveToFile(result, FileFormat.Docx)

			' Dispose of the document object
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
