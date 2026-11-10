Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shapes.Charts
Imports Spire.Doc.Interface

Namespace ReplaceTextIgnoreOfficeMath
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Initialize a new Document object to work with Word documents.
			Dim document As New Document()

			' Load an existing Word document from the specified file path into the document object.
			document.LoadFromFile("..\..\..\..\..\..\Data\ReplaceTextIgnoreOfficeMath.docx")

			' Create a new SearchOptions instance to customize search behavior during text replacement.
			Dim searchOptions As New SearchOptions()

			' Configure the search to skip Office Math objects (e.g., equations) during replacement.
			searchOptions.IgnoreOfficeMath = True

			' Ensure that only whole words matching the search term are replaced (not partial matches).
			searchOptions.FindWholeWordsOnly = True

			' Replace all occurrences of the word "a" with "an" in the document, using the defined search options.
			document.Replace("a", "an", searchOptions)

			' Save the modified document to a file in Word 2016 format.
			document.SaveToFile("ReplaceTextIgnoreOfficeMath_result.docx", FileFormat.Docx2016)

			' Close the document to release internal resources such as file handles and memory buffers.
			document.Close()

			' Explicitly dispose of the document object to immediately free unmanaged resources and memory.
			document.Dispose()


			'Launching the Word file.
			WordDocViewer("ReplaceTextIgnoreOfficeMath_result.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
