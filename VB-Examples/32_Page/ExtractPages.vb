Imports Spire.Doc

Namespace ExtractPages
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class to represent a Word document in memory.
			Dim document As New Document()

			' Load content from an existing .docx file located at the specified relative path into the Document object.
			document.LoadFromFile("..\..\..\..\..\..\Data\MultiplePages.docx")

			' Extract pages 0 through 2 from the original document and return them as a new Document object.
			Dim extractPage As Document = document.ExtractPages(0, 2)

			' Save the extracted pages to a new file.
			extractPage.SaveToFile("ExtractPages.docx", FileFormat.Docx)

			' Close the document to release any file handles or internal resources associated with it.
			document.Close()

			' Immediately free unmanaged resources used by the document object by calling Dispose (good practice for memory management).
			document.Dispose()

			'view the PDF file.
			WordDocViewer("ExtractPages.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
