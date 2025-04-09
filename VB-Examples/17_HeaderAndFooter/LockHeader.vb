Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace LockHeader
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\HeaderAndFooter.docx"

			'Create a Word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the first section
			Dim section As Section = doc.Sections(0)

			'Protect the document and set the ProtectionType as AllowOnlyFormFields
			doc.Protect(ProtectionType.AllowOnlyFormFields, "123")

			'Set the ProtectForm as false to unprotect the section
			section.ProtectForm = False

			'Save the document
			Dim output As String = "LockHeader.docx"
			doc.SaveToFile(output, FileFormat.Docx)

			'Dispose the document
			doc.Dispose()
			
			Viewer(output)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
