Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace CreateBookmarkForTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document.
			Dim document As New Document()

			'Add a new section.
			Dim section As Section = document.AddSection()

			'Create bookmark for a table
			CreateBookmarkForTable(document, section)

			Dim result As String = "Output.docx"
			'Save the document.
			document.SaveToFile(result, FileFormat.Docx)

			'Dispose the document
			document.Dispose()
			
			'Launch the Word file.
			WordDocViewer(result)
		End Sub

		Private Sub CreateBookmarkForTable(ByVal doc As Document, ByVal section As Section)
			'Add a paragraph
			Dim paragraph As Paragraph = section.AddParagraph()

			'Append text for added paragraph
			Dim txtRange As TextRange = paragraph.AppendText("The following example demonstrates how to create bookmark for a table in a Word document.")

			'Set the font in italic
			txtRange.CharacterFormat.Italic = True

			'Append bookmark start
			paragraph.AppendBookmarkStart("CreateBookmark")

			'Append bookmark end
			paragraph.AppendBookmarkEnd("CreateBookmark")

			'Add table
			Dim table As Table = section.AddTable(True)

			'Set the number of rows and columns
			table.ResetCells(2, 2)

			'Append text for table cells
			Dim range As TextRange = table(0, 0).AddParagraph().AppendText("sampleA")
			range = table(0, 1).AddParagraph().AppendText("sampleB")
			range = table(1, 0).AddParagraph().AppendText("120")
			range = table(1, 1).AddParagraph().AppendText("260")

			'Get the bookmark by index.
			Dim bookmark As Bookmark = doc.Bookmarks(0)

			'Get the name of bookmark.
			Dim bookmarkName As String = bookmark.Name

			'Locate the bookmark by name.
			Dim navigator As New BookmarksNavigator(doc)
			navigator.MoveToBookmark(bookmarkName)

			'Add table to TextBodyPart
			Dim part As TextBodyPart = navigator.GetBookmarkContent()
			part.BodyItems.Add(table)

			'Replace bookmark cotent with table
			navigator.ReplaceBookmarkContent(part)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
