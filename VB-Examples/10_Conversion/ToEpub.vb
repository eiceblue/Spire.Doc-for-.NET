Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc

Namespace ToEpub
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim doc As New Document()

			' Load the document from the specified file path
			doc.LoadFromFile("..\..\..\..\..\..\..\Data\ToEpub.doc")

			' Define the output file name for the EPUB file
			Dim result As String = "result.epub"

			' Save the document to an EPUB file with the specified output file name
			doc.SaveToFile(result, FileFormat.EPub)

			' Dispose of the Document object to release resources
			doc.Dispose()

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
