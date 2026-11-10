Imports Spire.Doc

Namespace SetPageLayout
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class to represent a Word document in memory.
			Dim document As New Document()

			' Load content from an existing .docx file located at the specified relative path into the Document object.
			document.LoadFromFile("..\..\..\..\..\..\..\Data\MultiplePages.docx")

			' Instantiate a ToPdfParameterList object to customize PDF export options.
			Dim toPdfParameter As New ToPdfParameterList()

			' Configure the page layout for the output PDF to display two pages per spread, starting with the left page.
			toPdfParameter.PageLayout = PdfPageLayout.TwoPageLeft

			' Export and save the Word document as a PDF file, applying the previously defined PDF settings.
			document.SaveToFile("SetPageLayout.pdf", toPdfParameter)

			' Close the document to release any file handles or internal resources associated with it.
			document.Close()

			' Immediately free unmanaged resources used by the document object by calling Dispose (good practice for memory management).
			document.Dispose()

			'view the PDF file.
			WordDocViewer("SetPageLayout.pdf")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
