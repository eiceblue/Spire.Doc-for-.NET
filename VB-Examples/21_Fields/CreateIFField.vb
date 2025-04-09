Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Interface

Namespace CreateIFField
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

			' Create an IF field and add it to the paragraph
			CreateIfField(document, paragraph)

			' Set field name and value for mail merge
			Dim fieldName() As String = { "Count" }
			Dim fieldValue() As String = { "2" }

			' Execute the mail merge
			document.MailMerge.Execute(fieldName, fieldValue)

			' Enable field update after mail merge
			document.IsUpdateFields = True

			' Specify the file name for saving the document
			Dim result As String = "Result-CreateAnIFField.docx"

			' Save the document to a file
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose the document object
			document.Dispose()
			
			'Launch the file.
			WordDocViewer(result)
		End Sub

			' Method to create an IF field
		   Private Shared Sub CreateIfField(ByVal document As Document, ByVal paragraph As Paragraph)
				' Create a new IF field
				Dim ifField As New IfField(document)
				ifField.Type = FieldType.FieldIf
				ifField.Code = "IF "

				' Add the IF field to the paragraph
				paragraph.Items.Add(ifField)

				' Add the merge field and condition to the paragraph
				paragraph.AppendField("Count", FieldType.FieldMergeField)
				paragraph.AppendText(" > ")
				paragraph.AppendText("""100"" ")
				paragraph.AppendText("""Thanks"" ")
				paragraph.AppendText("""The minimum order is 100 units""")

				' Create the end mark of the IF field and add it to the paragraph
				Dim [end] As IParagraphBase = document.CreateParagraphItem(ParagraphItemType.FieldMark)
				TryCast([end], FieldMark).Type = FieldMarkType.FieldEnd
				paragraph.Items.Add([end])

				' Set the end mark of the IF field
				ifField.End = TryCast([end], FieldMark)

		   End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
