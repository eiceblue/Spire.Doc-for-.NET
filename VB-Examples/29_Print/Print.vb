Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace Print
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of Document
			Dim document As New Document()

			' Load the Word document from the specified template file
			document.LoadFromFile("..\..\..\..\..\..\Data\Template.docx")

			' Create a new PrintDialog
			Dim dialog As New PrintDialog()

			' Allow printing of the current page
			dialog.AllowCurrentPage = True

			' Allow printing of a range of pages
			dialog.AllowSomePages = True

			' Use the system's default print dialog for selecting printer settings
			dialog.UseEXDialog = True

			Try
				' Set the PrintDialog property of the document to the created PrintDialog
				document.PrintDialog = dialog

				' Set the PrintDocument property of the PrintDialog to the document's PrintDocument
				dialog.Document = document.PrintDocument

				' Print the document using the PrintDialog
				dialog.Document.Print()
			Catch ex As Exception
				MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
			End Try
			
			' Dispose of the document object when finished using it
			document.Dispose()
		End Sub
	End Class
End Namespace
