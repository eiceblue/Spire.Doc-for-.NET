Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ManagePagination
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object
			Dim document As New Document()

			'Load a Word document from a specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			'Get the first section of the document
			Dim sec As Section = document.Sections(0)

			'Get the fifth paragraph of the section
			Dim para As Paragraph = sec.Paragraphs(4)

			'Set the PageBreakBefore property of the paragraph to True
			para.Format.PageBreakBefore = True

			'Specify the output file name
			Dim result As String = "Result-ManagePagination.docx"

			'Save the modified document to the specified file format (Docx2013)
			document.SaveToFile(result, FileFormat.Docx2013)

			'Dispose of the Document object to release resources
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
