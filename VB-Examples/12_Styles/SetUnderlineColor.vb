Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetUnderlineColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document instance
			Dim document As Document = New Document()

			' Add a new section to the document
			Dim section As Section = document.AddSection()

			' Add a new paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Append text to the paragraph and get the TextRange object for formatting
			Dim textRange As TextRange = paragraph.AppendText("Welcome to evaluate Spire.Doc for .NET product.")

			' Set the underline style of the text to single underline
			textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single

			' Set the underline color of the text to red
			textRange.CharacterFormat.UnderlineColor = Color.Red

			' Define the file path and name for saving the document
			Dim filePath As String = "SetUnderlineColor.docx"

			' Save the document to the specified file path in DOCX format
			document.SaveToFile(filePath, FileFormat.Docx)

			' Release resources used by the document
			document.Dispose()

			WordDocViewer(filePath)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
