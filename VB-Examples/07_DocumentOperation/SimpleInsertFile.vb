Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports System.IO

Namespace SimpleInsertFile
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim doc As New Document()

			' Load a Word document from a specified file path
			doc.LoadFromFile("..\..\..\..\..\..\..\Data\Template_N5.docx")

			' Insert the text from another Word document into the current document
			doc.InsertTextFromFile("..\..\..\..\..\..\..\Data\Template_N3.docx", FileFormat.Auto)

			' Specify the output file name for saving the modified document
			Dim output As String = "SimpleInsertFile_out.docx"

			' Save the document to the specified output file path in the DOCX format (version: Word 2013)
			doc.SaveToFile(output, FileFormat.Docx2013)

			' Release resources used by the Document object
			doc.Dispose()

			'Launch the document
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
