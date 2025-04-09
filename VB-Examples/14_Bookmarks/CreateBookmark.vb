Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace CreateBookmark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a Document object
			Dim document As New Document()

			'Add a Section to the document
			Dim section As Section = document.AddSection()

			'Call the CreateBookmark method to create bookmarks in the section
			CreateBookmark(section)

			'Save the document.
			document.SaveToFile("Output.docx", FileFormat.Docx)

			'Dispose the document
			document.Dispose()

			'Launch the Word file.
			WordDocViewer("Output.docx")
		End Sub

		Private Sub CreateBookmark(ByVal section As Section)
			'Add a Paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			'Add text with formatting and make it italic
			Dim txtRange As TextRange = paragraph.AppendText("The following example demonstrates how to create bookmark in a Word document.")
			txtRange.CharacterFormat.Italic = True

			'Add an empty paragraph for spacing
			section.AddParagraph()

			'Add another paragraph with heading style and colored text
			paragraph = section.AddParagraph()
			txtRange = paragraph.AppendText("Simple Create Bookmark.")
			txtRange.CharacterFormat.TextColor = Color.CornflowerBlue
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			'Add an empty paragraph for spacing
			section.AddParagraph()

			'Add a paragraph with a bookmark
			paragraph = section.AddParagraph()
			paragraph.AppendBookmarkStart("SimpleCreateBookmark")
			paragraph.AppendText("This is a simple bookmark.")
			paragraph.AppendBookmarkEnd("SimpleCreateBookmark")

			'Add an empty paragraph for spacing
			section.AddParagraph()

			'Add another paragraph with heading style and colored text
			paragraph = section.AddParagraph()
			txtRange = paragraph.AppendText("Nested Create Bookmark.")
			txtRange.CharacterFormat.TextColor = Color.CornflowerBlue
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			'Add an empty paragraph for spacing
			section.AddParagraph()

			'Add a paragraph with nested bookmarks
			paragraph = section.AddParagraph()
			paragraph.AppendBookmarkStart("Root")
			txtRange = paragraph.AppendText(" This is Root data ")
			txtRange.CharacterFormat.Italic = True
			paragraph.AppendBookmarkStart("NestedLevel1")
			txtRange = paragraph.AppendText(" This is Nested Level1 ")
			txtRange.CharacterFormat.Italic = True
			txtRange.CharacterFormat.TextColor = Color.DarkSlateGray
			paragraph.AppendBookmarkStart("NestedLevel2")
			txtRange = paragraph.AppendText(" This is Nested Level2 ")
			txtRange.CharacterFormat.Italic = True
			txtRange.CharacterFormat.TextColor = Color.DimGray
			paragraph.AppendBookmarkEnd("NestedLevel2")
			paragraph.AppendBookmarkEnd("NestedLevel1")
			paragraph.AppendBookmarkEnd("Root")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
