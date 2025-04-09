Imports Spire.Doc

Namespace Replace
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object.
			Dim document As New Document()

			' Load the Word document from the specified file path.
			document.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			' Replace all occurrences of "word" with "ReplacedText" in the document.
			' The replacement is case-insensitive and replaces whole words only.
			document.Replace("word", "ReplacedText", False, True)

			' Save the modified document to the same file, overwriting the original file.
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			' Dispose of the Document object to release resources.
			document.Dispose()

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
