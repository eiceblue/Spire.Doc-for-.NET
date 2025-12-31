Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace InsertImageAtBookmark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\Bookmark.docx"

			'Create a Word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Create an instance of BookmarksNavigator
			Dim bn As New BookmarksNavigator(doc)

			'Find a bookmark named Test
			bn.MoveToBookmark("Test", True, True)

			'Add a section
			Dim section0 As Section = doc.AddSection()

			'Add a paragraph for the section
			Dim paragraph As Paragraph = section0.AddParagraph()

			'Load an image
			Dim image As Image = image.FromFile("..\..\..\..\..\..\Data\Word.png")

			'Add an image into the paragraph
			Dim picture As DocPicture = paragraph.AppendPicture(image)
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim picture As DocPicture = paragraph.AppendPicture("..\..\..\..\..\..\Data\Word.png")
			' =============================================================================

			'Add the paragraph at the position of bookmark
			bn.InsertParagraph(paragraph)

			'Remove the section0
			doc.Sections.Remove(section0)

			'Save and launch document
			Dim output As String = "InsertImageAtBookmark.docx"
			doc.SaveToFile(output, FileFormat.Docx)

			'Dispose the document
			doc.Dispose()
			Viewer(output)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
