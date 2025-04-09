Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SetBeforOrAfterSpacingLines
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim doc As New Document()

			' Load a Word document from a specific file path
			doc.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			' Access the first section of the document
			Dim section As Section = doc.Sections(0)

			' Access the first paragraph in the section
			Dim paragraph As Paragraph = section.Paragraphs(0)

			' Set the spacing before the paragraph 
			paragraph.Format.BeforeSpacingLines = 5f

			' Set the spacing after the paragraph
			paragraph.Format.AfterSpacingLines = 15f

			' Save the modified document to a new file
			doc.SaveToFile("setBeforOrAfterSpacingLines.docx")

			' Dispose of the Document object to release resources
			doc.Dispose()

			WordDocViewer("setBeforOrAfterSpacingLines.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch e As Exception
				Debug.Write(e.StackTrace)
			End Try
		End Sub

	End Class
End Namespace
