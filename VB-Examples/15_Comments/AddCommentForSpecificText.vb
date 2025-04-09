Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AddCommentForSpecificText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a Word document
			Dim document As New Document()

			'Load the file from disk
			document.LoadFromFile("..\..\..\..\..\..\Data\CommentTemplate.docx")

			'Insert comments
			InsertComments(document, "development")

			'Save the document
			document.SaveToFile("AddCommentForTextRange.docx", FileFormat.Docx)

			'Dispose the document
			document.Dispose()

			'Launch the Word file.
			WordDocViewer("AddCommentForTextRange.docx")
		End Sub

		Private Sub InsertComments(ByVal doc As Document, ByVal keystring As String)
			'Find the key string
			Dim find As TextSelection = doc.FindString(keystring, False, True)

			'Create the commentmarkStart and commentmarkEnd
			Dim commentmarkStart As New CommentMark(doc)
			commentmarkStart.Type = CommentMarkType.CommentStart
			commentmarkStart.CommentId = 1
			Dim commentmarkEnd As New CommentMark(doc)
			commentmarkEnd.Type = CommentMarkType.CommentEnd
			commentmarkEnd.CommentId = 1

			'Add the content for comment
			Dim comment As New Comment(doc)

			'Add the text to the paragraph
			comment.Body.AddParagraph().Text = "Test comments"

			'Add author information
			comment.Format.Author = "E-iceblue"

			'Get the textRange
			Dim range As TextRange = find.GetAsOneRange()

			'Get its paragraph
			Dim para As Paragraph = range.OwnerParagraph

			'Get the index of textRange 
			Dim index As Integer = para.ChildObjects.IndexOf(range)

			'Add comment
			para.ChildObjects.Add(comment)

			'Insert the commentmarkStart and commentmarkEnd
			para.ChildObjects.Insert(index, commentmarkStart)
			para.ChildObjects.Insert(index + 2, commentmarkEnd)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
