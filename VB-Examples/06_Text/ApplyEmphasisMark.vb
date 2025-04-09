Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ApplyEmphasisMark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object.
			Dim document As New Document()

			'Load a Word document from a specified file location.
			document.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			'Find all occurrences of the text "Spire.Doc for .NET" in the document and store them in an array of TextSelection objects.
			Dim textSelections() As TextSelection = document.FindAllString("Spire.Doc for .NET", False, True)

			'Iterate through each TextSelection object in the array.
			For Each selection As TextSelection In textSelections
				'Set the EmphasisMark property of the CharacterFormat object associated with the selected text to Emphasis.Dot.
				selection.GetAsOneRange().CharacterFormat.EmphasisMark = Emphasis.Dot
			Next selection

			'Specify the output file name for the modified document.
			Dim output As String = "ApplyEmphasisMark.docx"

			'Save the modified document to the specified output file location.
			document.SaveToFile(output, FileFormat.Docx)

			'Dispose of the Document object to release system resources.
			document.Dispose()

			'Launching the file
			WordDocViewer(output)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
