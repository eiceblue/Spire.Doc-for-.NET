Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace PageFormat
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
			doc.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			' Move the navigator's cursor to the first section (section index 0) of the document.
			navigator.MoveToSection(0)

			' Set the page margins for the current section.
			navigator.PageSetup.Margins = New MarginsF(100, 80, 100, 80)

			' Set the page size of the current section to Letter.
			navigator.PageSetup.PageSize = PageSize.Letter

			' Save the modified document to a new file.
			doc.SaveToFile("PageFormat.docx", FileFormat.Docx)

			' Close the document to release internal resources.
			doc.Close()

			' Explicitly dispose of the document object to free memory immediately.
			doc.Dispose()

			'Launch the MS Word file.
			WordDocViewer("PageFormat.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
