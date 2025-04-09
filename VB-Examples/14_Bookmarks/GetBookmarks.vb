Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.Text
Imports System.IO

Namespace GetBookmarks
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
			Dim document As New Document()

			'Load the document from disk.
			document.LoadFromFile("..\..\..\..\..\..\Data\Bookmarks.docx")

			'Get the bookmark by index.
			Dim bookmark1 As Bookmark = document.Bookmarks(0)

			'Get the bookmark by name.
			Dim bookmark2 As Bookmark = document.Bookmarks("Test2")

			'Create StringBuilder to save 
			Dim content As New StringBuilder()

			'Set string format for displaying
			Dim result As String = String.Format("The bookmark obtained by index is " & bookmark1.Name & "." & vbCrLf & "The bookmark obtained by name is " & bookmark2.Name & "." & vbLf)

			'Add result string to StringBuilder
			content.AppendLine(result)

			'Save to a txt file
			File.WriteAllText("Bookmarks.txt", content.ToString())

			'Dispose the document
			document.Dispose()
			
			'Launch the file
			FileViewer("Bookmarks.txt")
		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
