Imports System.ComponentModel
Imports System.IO
Imports System.Net
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace DownloadWordFromURL
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Create a new instance of WebClient
			Dim webClient As New WebClient()

			' Download the Word file from the specified URL and load it into a memory stream
			Using ms As New MemoryStream(webClient.DownloadData("http://www.e-iceblue.com/images/test.docx"))
				document.LoadFromStream(ms, FileFormat.Docx)
			End Using

			' Set the file name for the result
			Dim result As String = "Result-DownloadWordFileFromURL.docx"

			' Save the loaded document to a new file in Docx 2013 format
			document.SaveToFile(result, FileFormat.Docx2013)

			' Release all resources used by the Document object
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
