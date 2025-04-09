Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace InsertPageBreakSecondApproach
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load an existing document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			' Append a page break at the end of the fourth paragraph in the first section of the document
			document.Sections(0).Paragraphs(3).AppendBreak(BreakType.PageBreak)

			' Specify the file name for the resulting document
			Dim result As String = "Result-InsertWordPageBreak.docx"

			' Save the modified document to a new file with the specified file format
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document object
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
