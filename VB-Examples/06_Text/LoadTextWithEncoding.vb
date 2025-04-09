Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace LoadTextWithEncoding
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Load the file path of the input file
			Dim inputFile As String = "../../../../../../Data/Sample_UTF-7.txt"

			'Create a new instance of the Document class
			Dim document As New Document()

			'Load the text from the input file using UTF-7 encoding
			document.LoadText(inputFile, Encoding.UTF7)

			'Specify the file path for the resulting file
			Dim resultFile As String = "LoadTextWithEncoding_out.docx"

			'Save the document to a Word file in the specified format (Docx)
			document.SaveToFile(resultFile, FileFormat.Docx)

			'Dispose of the document object to release any associated resources
			document.Dispose()

			WordDocViewer(resultFile)

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
