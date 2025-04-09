Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace AllowLatinTextWrapInMiddleOfAWord
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new Document object called document
			Dim document As New Document()

			'Load a Word document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\AllowLatinTextWrapInMiddleOfAWord.docx")

			'Get the first paragraph from the first section of the document and assign it to para
			Dim para As Paragraph = document.Sections(0).Paragraphs(0)

			'Set the WordWrap property of para's format to False, allowing Latin text to wrap in the middle of a word
			para.Format.WordWrap = False

			'Specify the output file name
			Dim result As String = "AllowLatinTextWrapInMiddleOfAWord-Result.docx"

			'Save the modified document as a Word document with the specified file format
			document.SaveToFile(result, FileFormat.Docx2013)

			'Dispose of the document object to release resources
			document.Dispose()

			'Launching the Word file.
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
