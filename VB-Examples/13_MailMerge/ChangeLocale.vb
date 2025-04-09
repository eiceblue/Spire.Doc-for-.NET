Imports System.Globalization
Imports System.Threading
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ChangeLocale
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

			'Store the current culture so it can be set back once mail merge is complete.
			Dim currentCulture As CultureInfo = Thread.CurrentThread.CurrentCulture

			'Set the current thread culture
			Thread.CurrentThread.CurrentCulture = New CultureInfo("de-DE")

			'Define the field values for the mail merge
			Dim fieldNames() As String = {"Contact Name", "Fax", "Date"}
			Dim fieldValues() As String = {"John Smith", "+1 (69) 123456", Date.Now.ToString()}

			'excute mail merge
			document.MailMerge.Execute(fieldNames, fieldValues)

			'restore the thread culture
			Thread.CurrentThread.CurrentCulture = currentCulture

			'Save doc file.
			Dim output As String = "ChangeLocale.docx"
			document.SaveToFile(output, FileFormat.Docx)

			'Dispose the document
			document.Dispose()

			'Launching the Word file.
			WordDocViewer(output)


		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
