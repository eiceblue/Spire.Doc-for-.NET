Imports Spire.Doc

Namespace AddHeaderOnlyFirstPage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\HeaderAndFooter.docx"

			'Create a Word document
			Dim doc1 As New Document()

			'Load the source file
			doc1.LoadFromFile(input)

			'Get the header from the first section
			Dim header As HeaderFooter = doc1.Sections(0).HeadersFooters.Header

			input = "..\..\..\..\..\..\Data\MultiplePages.docx"

			'Create another Word document
			Dim doc2 As New Document()

			'Load the destination file
			doc2.LoadFromFile(input)

			'Get the first page header of the destination document
			Dim firstPageHeader As HeaderFooter = doc2.Sections(0).HeadersFooters.FirstPageHeader

			'Loop the sections of doc2
			For Each section As Section In doc2.Sections

				'Specify that the current section has a different header/footer for the first page
				section.PageSetup.DifferentFirstPageHeaderFooter = True
			Next section

			'Removes all child objects in firstPageHeader
			firstPageHeader.Paragraphs.Clear()

			'Loop through the child objects of the header
			For Each obj As DocumentObject In header.ChildObjects

				'Add all child objects of the header to firstPageHeader
				firstPageHeader.ChildObjects.Add(obj.Clone())
			Next obj

			'Save the file
			Dim resultfile As String = "AddHeaderOnlyFirstPage.docx"
			doc2.SaveToFile(resultfile, FileFormat.Docx)

			'Dispose the document
			doc1.Dispose()
			doc2.Dispose()
			
			Viewer(resultfile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
