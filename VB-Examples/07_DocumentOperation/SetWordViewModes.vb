Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SetWordViewModes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the document from the specified file
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			' Set the document view type, zoom percent, and zoom type for the ViewSetup
			document.ViewSetup.DocumentViewType = DocumentViewType.WebLayout
			document.ViewSetup.ZoomPercent = 150
			document.ViewSetup.ZoomType = ZoomType.None

			' Set the file name for the result
			Dim result As String = "Result-SetWordViewModes.docx"

			' Save the modified document to a new file in Docx 2013 format
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
