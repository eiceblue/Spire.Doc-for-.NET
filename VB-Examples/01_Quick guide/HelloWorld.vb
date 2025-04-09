Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace HelloWorld
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Add a Section to the document
			Dim section As Section = document.AddSection()

			' Add a Paragraph to the Section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Append text to the paragraph
			paragraph.AppendText("Hello World!")

			' Save the document to a file named "Sample.docx" in Docx format
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			' Dispose of the Document object to release resources
			document.Dispose()

			'Launching the Word file.
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
