Imports Spire.Doc

Namespace MailMerage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a Word document
			Dim document As New Document()

			'Load the file from disk
			document.LoadFromFile("..\..\..\..\..\..\Data\MailMerge.doc")

			'Define the field names for the mail merge
			Dim fieldNames() As String = {"Contact Name", "Fax", "Date"}

			'Define the field values for the mail merge
			Dim fieldValues() As String = {"John Smith", "+1 (69) 123456", Date.Now.Date.ToString()}

			'Begin the mail merge process
			document.MailMerge.Execute(fieldNames, fieldValues)

			'Save as Doc file.
			document.SaveToFile("Sample.doc", FileFormat.Doc)

			'Dispose the document
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("Sample.doc")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
