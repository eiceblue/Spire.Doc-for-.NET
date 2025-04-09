Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc

Namespace RemoveReadOnlyRestriction
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the Word document file from the specified path
			document.LoadFromFile("..\..\..\..\..\..\Data\RemoveReadOnlyRestriction.docx")

			' Remove the read-only restriction from the document
			document.Protect(ProtectionType.NoProtection)

			' Specify the output file name for the modified document
			Dim result As String = "RemoveReadOnlyRestriction_out.docx"

			' Save the modified document to the specified file format
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose the Document object to free resources
			document.Dispose()

			'Launch the file.
			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
