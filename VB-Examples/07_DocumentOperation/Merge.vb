Imports Spire.Doc

Namespace Merge
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
				' Create a new Document object
				Dim document As New Document()
				' Load a Word document from the specified file path
				document.LoadFromFile("..\..\..\..\..\..\..\Data\Summary_of_Science.doc", FileFormat.Doc)

				' Create a new Document object for merging
				Dim documentMerge As New Document()
				' Load another Word document for merging from the specified file path
				documentMerge.LoadFromFile("..\..\..\..\..\..\..\Data\Bookmark.docx", FileFormat.Docx)

				' Iterate through each Section in the documentMerge and add clones to the original document
				For Each sec As Section In documentMerge.Sections
					document.Sections.Add(sec.Clone())
				Next sec

				' Save the merged document to a new file with the specified file format
				document.SaveToFile("Sample.docx", FileFormat.Docx)

				' Dispose of the resources used by the Document objects
				document.Dispose()
				documentMerge.Dispose()

				'Launching the MS Word file.
				WordDocViewer("Sample.docx")


		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
