Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting
Imports System.IO
Imports Spire.Doc.Fields
Namespace LoadAndSaveToDisk
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Set the input file path
			Dim input As String = "..\..\..\..\..\..\Data\Template.docx"

			' Create a new Document object
			Dim doc As New Document()

			' Load the document from the specified input file
			doc.LoadFromFile(input)

			' Set the file name for the result
			Dim result As String = "LoadAndSaveToDisk_out.docx"

			' Save the loaded document to a new file in Docx format
			doc.SaveToFile(result, FileFormat.Docx)

			' Release all resources used by the Document object
			doc.Dispose()

			WordViewer(result)
		End Sub
		Private Sub WordViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
