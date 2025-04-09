Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace DetermineDocumentEncryption
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim isEncrypted As Boolean = Document.IsEncrypted("..\..\..\..\..\..\Data\TemplateWithPassword.docx")
			If isEncrypted = True Then
				MessageBox.Show("This document is encrypted. ")
			Else
				MessageBox.Show("This document is unencrypted. ")
			End If
		End Sub

	End Class
End Namespace
