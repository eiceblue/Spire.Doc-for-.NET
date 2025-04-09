Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ExtractBookmarkText
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

			'Creates a BookmarkNavigator instance to access the bookmark
			Dim navigator As New BookmarksNavigator(doc)

			'Locate a specific bookmark by bookmark name
			navigator.MoveToBookmark("Content")

			'Get the bookmark content
			Dim textBodyPart As TextBodyPart = navigator.GetBookmarkContent()

			'Iterate through the items in the bookmark content to get the text
			Dim text As String = Nothing
			For Each item In textBodyPart.BodyItems
				If TypeOf item Is Paragraph Then

					'Iterate through the child objects of the paragraph
					For Each childObject In (TryCast(item, Paragraph)).ChildObjects
						If TypeOf childObject Is TextRange Then
							'Append the text
							text &= (TryCast(childObject, TextRange)).Text
						End If
					Next childObject
				End If
			Next item

			'Save to TXT File
			Dim output As String = "ExtractBookmarkText.txt"
			File.WriteAllText(output, text)

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
