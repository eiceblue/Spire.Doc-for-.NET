Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RemoveImageWatermark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim document As New Document()

			' Load the document from a file
			document.LoadFromFile("..\..\..\..\..\..\Data\RemoveImageWatermark.docx")

			' Remove the watermark from the document
			document.Watermark = Nothing

			' Specify the output file name
			Dim result As String = "Result-RemoveImageWatermark.docx"

			' Save the modified document to a new file in Docx2013 format
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose the document object
			document.Dispose()

			'Launch the MS Word file.
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
