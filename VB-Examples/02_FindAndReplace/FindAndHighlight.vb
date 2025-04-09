Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace FindAndHighlight
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load a Word document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			' Find all occurrences of the string "word" in the document (case-insensitive)
			Dim textSelections() As TextSelection = document.FindAllString("word", False, True)

			' Iterate through each TextSelection
			For Each selection As TextSelection In textSelections
				' Set the highlight color of the selected text to yellow
				selection.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow
			Next selection

			' Save the modified document to a file named "Sample.docx"
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			' Dispose of the Document object to release resources
			document.Dispose()

			'Launching the  Word file.
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
