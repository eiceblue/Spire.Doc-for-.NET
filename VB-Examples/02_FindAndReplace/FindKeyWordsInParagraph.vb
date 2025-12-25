Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace FindKeyWordsInParagraph
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Input file path
			Dim input As String = "..\..\..\..\..\..\Data\Sample.docx"

			'Output file path
			Dim output As String = "FindKeyWordsInParagraph_output.docx"

			'Create word document
			Dim document As New Document()

			'Load a document
			document.LoadFromFile(input)

			'Get the first section
			Dim s As Section = document.Sections(0)

			'Get the second paragraph
			Dim para As Paragraph = s.Paragraphs(1)

			'Find all matched keywords
			Dim textSelections() As TextSelection = para.FindAllString("Word", False, True)

			'Highlight text
			For Each selection As TextSelection In textSelections
				selection.GetAsOneRange().CharacterFormat.HighlightColor= Color.FromArgb(255, 255, 0)
			Next selection

			' Save to file
			document.SaveToFile(output, FileFormat.Docx2019)

			'Dispose the document
			document.Dispose()

			'Launching the Word file.
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
