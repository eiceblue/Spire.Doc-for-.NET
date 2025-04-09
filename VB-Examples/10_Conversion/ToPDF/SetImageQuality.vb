Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SetImageQuality
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load a Word document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\..\Data\Template_Doc_1.doc", FileFormat.Doc)

			' Set the JPEG quality for the document to 40
			document.JPEGQuality = 40

			' Specify the file name for the resulting PDF
			Dim result As String = "Result-DocToPDFImageQuality.pdf"

			' Save the document to a PDF file with the specified image quality
			document.SaveToFile(result, FileFormat.PDF)

			' Dispose of the Document object to release resources
			document.Dispose()

			'Launch the Pdf file.
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
