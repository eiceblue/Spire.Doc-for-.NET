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
Namespace MailMergeFormField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			Dim input As String = "..\..\..\..\..\..\Data\MailMergeFormField.doc"

			'Create a Word document
			Dim document As New Document()

			'Load the document from the specified file
			document.LoadFromFile(input)

			'Define the field names for the mail merge
			Dim fieldNames() As String = {"Contact Name", "Fax", "Date", "Urgent", "Share", "Submit", "Body"}

			'Define the field values for the mail merge
			Dim fieldValues() As String = {"John Smith", "+1 (69) 123456", Date.Now.Date.ToString(), "Yes", "No", "Yes", "<b>It's very urgent. Please deal with it ASAP. </b>"}

			'Subscribe to the MergeField event
			AddHandler document.MailMerge.MergeField, AddressOf MailMerge_MergeField

			'Execute the mail merge using the field names and values
			document.MailMerge.Execute(fieldNames, fieldValues)

			'Save the merged document
			Dim result As String = "MailMergeFormField_out.docx"
			document.SaveToFile(result, FileFormat.Docx)

			'Dispose the document
			document.Dispose()
			
			WordViewer(result)
		End Sub

		Private Sub MailMerge_MergeField(ByVal sender As Object, ByVal args As MergeFieldEventArgs)
			If args.FieldValue.ToString = "Yes" Then

				'Get the checkbox name from the field name
				Dim checkBoxName As String = args.FieldName

				'Get the owner paragraph of the current merge field
				Dim para As Paragraph = args.CurrentMergeField.OwnerParagraph

				'Get the index of the current merge field within its parent paragraph
				Dim index As Integer = para.ChildObjects.IndexOf(args.CurrentMergeField)
				'Create a new CheckBoxFormField
				Dim field As CheckBoxFormField = TryCast(para.AppendField(checkBoxName, FieldType.FieldFormCheckBox), CheckBoxFormField)

				'Insert the new checkbox field at the same index as the current merge field
				para.ChildObjects.Insert(index, field)

				'Remove the current merge field from the paragraph
				para.ChildObjects.Remove(args.CurrentMergeField)

				'Set the checkbox field as checked
				field.Checked = True
			End If
			If args.FieldValue.ToString = "No" Then

				'Get the checkbox name from the field name
				Dim checkBoxName As String = args.FieldName

				'Get the owner paragraph of the current merge field
				Dim para As Paragraph = args.CurrentMergeField.OwnerParagraph

				'Get the index of the current merge field within its parent paragraph
				Dim index As Integer = para.ChildObjects.IndexOf(args.CurrentMergeField)

				' Create a new CheckBoxFormField
				Dim field As CheckBoxFormField = TryCast(para.AppendField(checkBoxName, FieldType.FieldFormCheckBox), CheckBoxFormField)

				'Insert the new checkbox field at the same index as the current merge field
				para.ChildObjects.Insert(index, field)

				'Remove the current merge field from the paragraph
				para.ChildObjects.Remove(args.CurrentMergeField)

				'Set the checkbox field as unchecked
				field.Checked = False
			End If

			If args.FieldName = "Body" Then

				' Insert text input form field.
				Dim para As Paragraph = args.CurrentMergeField.OwnerParagraph

				'Append the HTML content as plain text to the paragraph
				para.AppendHTML(args.FieldValue.ToString())

				'Remove the current merge field from the paragraph
				para.ChildObjects.Remove(args.CurrentMergeField)
			End If

			If args.FieldName = "Date" Then

				'Get the text input name from the field name
				Dim textInputName As String = args.FieldName

				'Get the owner paragraph of the current merge field
				Dim para As Paragraph = args.CurrentMergeField.OwnerParagraph

				'Create a new TextFormField
				Dim field As TextFormField = TryCast(para.AppendField(textInputName, FieldType.FieldFormTextInput), TextFormField)

				'Remove the current merge field from the paragraph
				para.ChildObjects.Remove(args.CurrentMergeField)

				'Set the text value for the text input field
				field.Text = args.FieldValue.ToString()
			End If
		End Sub

		Private Sub WordViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
