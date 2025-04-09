Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO
Imports Spire.Doc.Fields
Imports System.Data.OleDb
Namespace MailMergeSwitches
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\MailMergeSwitches.docx"

			'Create a Word document
			Dim doc As New Document()

			'Load a mail merge template file
			doc.LoadFromFile(input)

			'Define the field names for the mail merge
			Dim fieldName() As String = {"XX_Name"}

			'Define the field values for the mail merge
			Dim fieldValue() As String = {"Jason Tang"}

			'Execute the mail merge using the field names and values
			doc.MailMerge.Execute(fieldName, fieldValue)
			
			'Save to file
			Dim result As String = "MailMergeSwitches_out.docx"
			doc.SaveToFile(result, FileFormat.Docx)

			'Dispose the document
			doc.Dispose()
			
			WordViewer(result)
		End Sub
		Private Sub WordViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
