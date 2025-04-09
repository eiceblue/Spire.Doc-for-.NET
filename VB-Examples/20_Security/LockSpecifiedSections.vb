Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace LockSpecifiedSections
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Add two sections to the document
			Dim s1 As Section = document.AddSection()
			Dim s2 As Section = document.AddSection()

			' Add a paragraph with text to section 1
			s1.AddParagraph().AppendText("Spire.Doc demo, section 1")

			' Add a paragraph with text to section 2
			s2.AddParagraph().AppendText("Spire.Doc demo, section 2")

			' Protect the document with a password and allow only form fields
			document.Protect(ProtectionType.AllowOnlyFormFields, "123")

			' Disable form field protection for section 2
			s2.ProtectForm = False

			' Specify the output file path for the locked document
			Dim result As String = "Result-LockSpecifiedSections.docx"

			' Save the locked document to the output file path in DOCX format (compatible with Word 2013)
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose the document object to free up resources
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
