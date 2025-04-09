Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.Drawing.Printing

Namespace CallPreviewControl
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path
			Dim input As String = "..\..\..\..\..\..\Data\Sample.docx"

			' Create a new instance of Document
			Dim doc As New Document()

			' Load the Word document from the specified input file
			doc.LoadFromFile(input)

			' Get the PrintDocument associated with the document
			Dim printDoc As PrintDocument = doc.PrintDocument

			' Create a new PrintPreviewDialog
			Dim printPreviewDialog As New PrintPreviewDialog()

			' Set the PrintDocument for the PrintPreviewDialog
			printPreviewDialog.Document = doc.PrintDocument

			' Set the size of the PrintPreviewDialog's client area
			printPreviewDialog.ClientSize = New Size(600, 800)

			' Show the PrintPreviewDialog
			printPreviewDialog.ShowDialog()

			' Dispose of the document object when finished using it
			doc.Dispose()
		End Sub
	End Class
End Namespace
