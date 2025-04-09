Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetEditableRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the Word document file from the specified path
			document.LoadFromFile("..\..\..\..\..\..\Data\SetEditableRange.docx")

			' Set the document protection to allow only reading with a password
			document.Protect(ProtectionType.AllowOnlyReading, "password")

			' Create a PermissionStart object to mark the start of an editable range with a specific ID
			Dim start As New PermissionStart(document, "testID")
			' Create a PermissionEnd object to mark the end of the editable range with the same ID
			Dim [end] As New PermissionEnd(document, "testID")

			' Insert the PermissionStart object at the beginning of the first paragraph in the first section
			document.Sections(0).Paragraphs(0).ChildObjects.Insert(0, start)
			' Add the PermissionEnd object to the end of the first paragraph in the first section
			document.Sections(0).Paragraphs(0).ChildObjects.Add([end])

			' Specify the output file name for the modified document
			Dim output As String = "SetEditableRange_output.docx"

			' Save the modified document to the specified file format
			document.SaveToFile(output, FileFormat.Docx)

			' Dispose the Document object to free resources
			document.Dispose()
			
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
