Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO

Namespace ConvertEpubToPDF
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As Document = New Document()

			' Load a Word document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\..\Data\ToPDF.epub", FileFormat.EPub)

			' Save the document as a PDF file with the name "Sample.pdf"
			document.SaveToFile("ConvertEpubToPDF.pdf", FileFormat.PDF)

			' Dispose of the Document object to free up resources
			document.Dispose()
			FileViewer("ConvertEpubToPDF.pdf")
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
