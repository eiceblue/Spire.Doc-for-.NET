Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace MoveToHeaderAndFooter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new empty document instance.
			Dim doc As New Document()

			' Create a document navigator to help navigate and modify the document content.
			Dim navigator As New DocumentNavigator(doc)

			' Load an existing Word document from the specified relative file path.
			doc.LoadFromFile("..\..\..\..\..\..\Data\MoveToHeaderAndFooter.docx")

			' Move the navigator's cursor to the first section of the document (section index 0).
			navigator.MoveToSection(0)

			' Navigate to the footer of the first page in the current section.
			navigator.MoveToHeaderFooter(HeaderFooterType.FooterFirstPage)

			' Write a new line of text into the first-page footer.
			navigator.Writeln("The footer on the first page.")

			' Navigate to the header of the first page in the current section.
			navigator.MoveToHeaderFooter(HeaderFooterType.HeaderFirstPage)

			' Write a new line of text into the first-page header.
			navigator.Writeln("The header on the first page.")

			' Save the modified document to a new file named "MoveToHeaderAndFooter.docx" in DOCX format.
			doc.SaveToFile("MoveToHeaderAndFooter.docx", FileFormat.Docx)

			' Close the document to release internal resources.
			doc.Close()

			' Explicitly dispose of the document object to free memory immediately.
			doc.Dispose()

			'Launch the MS Word file.
			WordDocViewer("MoveToHeaderAndFooter.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
