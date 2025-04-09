Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO
Imports Spire.Doc.Fields
Imports System.Data.OleDb
Imports System.Linq
Imports Spire.Doc.Reporting
Imports System.Collections
Imports Spire.Doc.Interface
Namespace ExecuteConditionalField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object
			Dim doc As New Document()

			'Add a new section 
			Dim section As Section = doc.AddSection()

			'Add a new paragraph to the section 
			Dim paragraph As Paragraph = section.AddParagraph()

			'Create and add the first IF field
			CreateIFField1(doc, paragraph)

			'Add another paragraph to the section
			paragraph = section.AddParagraph()

			'Create and add the second IF field
			CreateIFField2(doc, paragraph)

			'Define the field names for the mail merge
			Dim fieldName() As String = {"Count", "Age"}

			'Define the field values for the mail merge
			Dim fieldValue() As String = {"2", "30"}

			'Execute the mail merge
			doc.MailMerge.Execute(fieldName, fieldValue)

			'Set IsUpdateFields property to true
			doc.IsUpdateFields = True

			'Save the document 
			doc.SaveToFile("ExecuteConditionalField_result.docx", FileFormat.Docx)

			'Dispose the document
			doc.Dispose()
			WordViewer("ExecuteConditionalField_result.docx")
		End Sub
		Private Sub CreateIFField1(ByVal document As Document, ByVal paragraph As Paragraph)
			'Create a new IfField object
			Dim ifField As New IfField(document)

			'Set the type and code of the IfField
			ifField.Type = FieldType.FieldIf
			ifField.Code = "IF "

			'Add the IfField to the paragraph
			paragraph.Items.Add(ifField)

			'Append the fields and text to the paragraph
			paragraph.AppendField("Count", FieldType.FieldMergeField)
			paragraph.AppendText(" > ")
			paragraph.AppendText("""1"" ")
			paragraph.AppendText("""Greater than one"" ")
			paragraph.AppendText("""Less than one""")

			'Create and add the end field mark
			Dim [end] As IParagraphBase = document.CreateParagraphItem(ParagraphItemType.FieldMark)
			TryCast([end], FieldMark).Type = FieldMarkType.FieldEnd
			paragraph.Items.Add([end])

			'Set the end field mark for the IfField
			ifField.End = TryCast([end], FieldMark)
		End Sub

		Private Sub CreateIFField2(ByVal document As Document, ByVal paragraph As Paragraph)
			'Create a new IfField object
			Dim ifField As New IfField(document)

			'Set the type and code of the IfField
			ifField.Type = FieldType.FieldIf
			ifField.Code = "IF "

			'Add the IfField to the paragraph
			paragraph.Items.Add(ifField)

			'Append the fields and text to the paragraph
			paragraph.AppendField("Age", FieldType.FieldMergeField)
			paragraph.AppendText(" > ")
			paragraph.AppendText("""50"" ")
			paragraph.AppendText("""The old man"" ")
			paragraph.AppendText("""The young man""")

			'Create and add the end field mark
			Dim [end] As IParagraphBase = document.CreateParagraphItem(ParagraphItemType.FieldMark)
			TryCast([end], FieldMark).Type = FieldMarkType.FieldEnd
			paragraph.Items.Add([end])

			'Set the end field mark for the IfField
			ifField.End = TryCast([end], FieldMark)
		End Sub
		Private Sub WordViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
