Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ListFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document instance.
			Dim doc As New Document()

			' Add a new section to the document.
			Dim section As Section = doc.AddSection()

			' Add the first paragraph to the section.
			Dim paragraph As Paragraph = section.AddParagraph()

			' Append text to the first paragraph.
			paragraph.AppendText("This is the first paragraph.")

			' Add a second paragraph to the section.
			paragraph = section.AddParagraph()

			' Append text to the second paragraph.
			paragraph.AppendText("This is the second paragraph.")

			' Add a third paragraph to the section.
			paragraph = section.AddParagraph()

			' Append text to the third paragraph.
			paragraph.AppendText("This is the third paragraph.")

			' Create a document navigator to facilitate navigation and formatting.
			Dim navigator As New DocumentNavigator(doc)

			' Apply bullet list style to the current position (first paragraph by default).
			navigator.ListFormat.ApplyBulletStyle()

			' Move the navigator's cursor to the third paragraph (index 2) at character offset 0.
			navigator.MoveToParagraph(2, 0)

			' Apply bullet list style to the third paragraph.
			navigator.ListFormat.ApplyBulletStyle()

			' Save the document to a file named "ListFormat.docx" in DOCX format.
			doc.SaveToFile("ListFormat.docx", FileFormat.Docx)

			' Close the document to release internal resources.
			doc.Close()

			' Explicitly dispose of the document object to free memory immediately.
			doc.Dispose()

			'Launch the MS Word file.
			WordDocViewer("ListFormat.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
