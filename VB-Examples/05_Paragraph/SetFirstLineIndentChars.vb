Imports System.ComponentModel
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetFirstLineIndentChars
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load a Word document from the specified path
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			' Create a new Paragraph object with the loaded document
			Dim para As New Paragraph(document)

			' Append a text range to the paragraph and set its properties
			Dim textRange1 As TextRange = para.AppendText("This is an inserted paragraph.")
			textRange1.CharacterFormat.TextColor = Color.Blue
			textRange1.CharacterFormat.FontSize = 15

			' Set the first-line indent of the paragraph to 2 characters
			para.Format.FirstLineIndentChars = 2

			' Uncomment the following line to set the hanging indent as 2 characters
			' para.Format.FirstLineIndentChars = -2;

			' Set the first-line indent of the paragraph to 0 characters
			para.Format.SetFirstLineIndentChars(0)

			' Insert the paragraph at index 1 in the first section of the document
			document.Sections(0).Paragraphs.Insert(1, para)

			' Specify the file name for the resulting document
			Dim result As String = "Result-SetFirstLineIndentChars.docx"

			' Save the modified document to a file in Docx2013 format
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose the document object to free resources
			document.Dispose()

			'Launch the file.
			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
