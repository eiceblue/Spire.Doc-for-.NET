Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace RemoveContentWithComment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a document
			Dim document As New Document()

			'Load the document from disk.
			document.LoadFromFile("..\..\..\..\..\..\Data\Comments.docx")

			'Get the first comment
			Dim comment As Comment = document.Comments(0)

			'Get the paragraph of obtained comment
			Dim para As Paragraph = comment.OwnerParagraph

			'Get index of the CommentMarkStart 
			Dim startIndex As Integer = para.ChildObjects.IndexOf(comment.CommentMarkStart)

			'Get index of the CommentMarkEnd
			Dim endIndex As Integer = para.ChildObjects.IndexOf(comment.CommentMarkEnd)

			'Create a list
			Dim list As New List(Of TextRange)()

			'Get TextRanges between the indexes
			For i As Integer = startIndex To endIndex - 1
				If TypeOf para.ChildObjects(i) Is TextRange Then

					'Add the text range
					list.Add(TryCast(para.ChildObjects(i), TextRange))
				End If
			Next i

			'Insert a new TextRange
			Dim textRange As New TextRange(document)

			'clear the text
			textRange.Text = Nothing

			'Insert the new textRange
			para.ChildObjects.Insert(endIndex, textRange)

			'Remove previous TextRanges
			For i As Integer = 0 To list.Count - 1
				para.ChildObjects.Remove(list(i))
			Next i

			Dim result As String = "Output.docx"
			'Save the document.
			document.SaveToFile(result, FileFormat.Docx)

			'Dispose the document
			document.Dispose()

			'Launch the Word file.
			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
