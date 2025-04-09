Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SetBookmarkColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\BookmarkTemplate.docx"

			'Create a word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Create an instance of ToPdfParameterList
			Dim toPdf As New ToPdfParameterList()

			'Set CreateWordBookmarks to true to use word bookmarks when create the bookmarks
			toPdf.CreateWordBookmarks = True

			'Set the title of word bookmarks
			toPdf.WordBookmarksTitle = "Changed bookmark"

			'Set the text color of word bookmarks
			toPdf.WordBookmarksColor = Color.Gray

			'Call the event document_BookmarkLayout when drawing a bookmark
			AddHandler doc.BookmarkLayout, AddressOf document_BookmarkLayout

			'Save the document
			Dim output As String = "SetBookmarkColor.pdf"
			doc.SaveToFile(output, toPdf)

			'Dispose the document
			doc.Dispose()
			
			Viewer(output)
		End Sub
		Private Shared Sub document_BookmarkLayout(ByVal sender As Object, ByVal args As Spire.Doc.Documents.Rendering.BookmarkLevelEventArgs)
			'set the different color for different levels of bookmarks
			If args.BookmarkLevel.Level = 2 Then
				args.BookmarkLevel.Color = Color.Red
				args.BookmarkLevel.Style = BookmarkTextStyle.Bold
			ElseIf args.BookmarkLevel.Level = 3 Then
				args.BookmarkLevel.Color = Color.Gray
				args.BookmarkLevel.Style = BookmarkTextStyle.Italic
			Else
				args.BookmarkLevel.Color = Color.Green
				args.BookmarkLevel.Style = BookmarkTextStyle.Regular
			End If
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
