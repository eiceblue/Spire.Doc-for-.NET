Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SpecifiedProtectionType
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the Word document file from the specified path
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_2.docx")

			' Set the document protection to allow only reading with the specified password
			document.Protect(ProtectionType.AllowOnlyReading, "123456")

			' Specify the output file name for the modified document
			Dim result As String = "Result-SpecifiedProtectionType.docx"

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
