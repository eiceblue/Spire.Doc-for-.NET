Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetCultureForField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document
			Dim document As New Document()

			' Add a section to the document
			Dim section As Section = document.AddSection()

			' Add a paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Add text to the paragraph
			paragraph.AppendText("Add Date Field: ")

			' Append a date field to the paragraph and set its format
			Dim field1 As Field = TryCast(paragraph.AppendField("Date1", FieldType.FieldDate), Field)
			field1.Code = "DATE  \@" & """yyyy\MM\dd"""

			' Add a new paragraph to the section
			Dim newParagraph As Paragraph = section.AddParagraph()

			' Add text to the new paragraph
			newParagraph.AppendText("Add Date Field with setting French Culture: ")

			' Append a date field to the new paragraph and set its format
			Dim field2 As Field = newParagraph.AppendField("""\@""dd MMMM yyyy", FieldType.FieldDate)
			field2.CharacterFormat.LocaleIdASCII = 1036

			' Enable automatic update of fields in the document
			document.IsUpdateFields = True

			' Save the document to a file
			document.SaveToFile("Output.docx", FileFormat.Docx)

			' Dispose of the document object
			document.Dispose()

			'Launch the Word file.
			WordDocViewer("Output.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
