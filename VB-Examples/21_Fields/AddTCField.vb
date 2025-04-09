Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AddTCField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Add a new section to the document
			Dim section As Section = document.AddSection()

			' Add a new paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Append a TC (Table of Contents) field to the paragraph with the specified entry text
			Dim field As Field = paragraph.AppendField("TC", FieldType.FieldTOCEntry)
			field.Code = "TC " & """Entry Text""" & " \f" & " t"

			' Save the document to a file with the specified file name and format (Docx)
			document.SaveToFile("AddTCField.docx", FileFormat.Docx)

			' Dispose the Document object to free resources
			document.Dispose()

			'Launch result file and please set "Show all formatting marks" to display the field 
			WordDocViewer("AddTCField.docx")
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
