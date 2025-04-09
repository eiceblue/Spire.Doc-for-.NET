Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc

Namespace PreserveWordBookmarks
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()
			' Load the document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\..\Data\Sample.doc")

			' Create a new ToPdfParameterList object
			Dim toPdf As New ToPdfParameterList()
			' Set the CreateWordBookmarks property to True to preserve Word bookmarks in the PDF
			toPdf.CreateWordBookmarks = True
			' Set the WordBookmarksTitle to specify the title of the bookmarks
			toPdf.WordBookmarksTitle = "Bookmark"
			' Set the WordBookmarksColor to specify the color of the bookmarks
			toPdf.WordBookmarksColor = Color.Gray

			' Add event handler for the BookmarkLayout event
			AddHandler document.BookmarkLayout, AddressOf document_BookmarkLayout

			' Save the document to a PDF file with the specified parameters
			document.SaveToFile("PreserveBookmarks.pdf", toPdf)
			' Dispose of the Document object to release resources
			document.Dispose()

			'Launch the file.
			FileViewer("PreserveBookmarks.pdf")
		End Sub
		
		' Event handler for the BookmarkLayout event
		Private Sub document_BookmarkLayout(ByVal sender As Object, ByVal args As Spire.Doc.Documents.Rendering.BookmarkLevelEventArgs)

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

		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
