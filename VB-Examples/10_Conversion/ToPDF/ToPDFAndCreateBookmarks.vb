Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ToPDFAndCreateBookmarks
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path for the Word document
			Dim inputFile As String = "../../../../../../../Data/BookmarkTemplate.docx"

			' Specify the output file path for the generated PDF with bookmarks
			Dim outFile As String = "ToPDFAndCreateBookmarks_out.pdf"

			' Create a new Document object
			Dim document As New Document()

			' Load a Word document from the specified file path
			document.LoadFromFile(inputFile)

			' Create a new ToPdfParameterList object to set parameters for PDF conversion
			Dim parames As New ToPdfParameterList()

			' Set the option to create Word bookmarks in the PDF
			parames.CreateWordBookmarks = True

			' Set the option to create Word bookmarks using headings (commented out)
			' parames.CreateWordBookmarksUsingHeadings = true;

			' Set the option to create Word bookmarks using headings (false as per the commented line)
			parames.CreateWordBookmarksUsingHeadings = False

			' Save the document to a PDF file with the specified bookmark creation settings
			document.SaveToFile(outFile, parames)

			' Dispose of the Document object to release resources
			document.Dispose()

			WordDocViewer(outFile)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
