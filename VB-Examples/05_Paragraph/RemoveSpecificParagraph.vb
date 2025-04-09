Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RemoveSpecificParagraph
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new instance of the Document class.
			Dim document As New Document()

			'Load a Word document from the specified file path.
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			'Remove the first paragraph from the first section of the document.
			document.Sections(0).Paragraphs.RemoveAt(0)

			'Specify the file name for the resulting document after removing the paragraph.
			Dim result As String = "Result-RemoveSpecificParagraph.docx"

			'Save the modified document to the specified file path in Docx2013 format.
			document.SaveToFile(result, FileFormat.Docx2013)

			'Dispose of the Document object to release resources.
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
