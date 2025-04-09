Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace FromCommentRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new instance of the Document class.
			Dim sourceDoc As New Document()

			'Load a Word document from a specified file path.
			sourceDoc.LoadFromFile("..\..\..\..\..\..\Data\Comments.docx")

			'Create a new instance of the Document class to store the destination document.
			Dim destinationDoc As New Document()

			'Add a new section to the destination document.
			Dim destinationSec As Section = destinationDoc.AddSection()

			'Get the first comment from the source document.
			Dim comment As Comment = sourceDoc.Comments(0)

			'Get the paragraph that owns the comment.
			Dim para As Paragraph = comment.OwnerParagraph

			'Get the index of the comment's start mark in the paragraph.
			Dim startIndex As Integer = para.ChildObjects.IndexOf(comment.CommentMarkStart)

			'Get the index of the comment's end mark in the paragraph.
			Dim endIndex As Integer = para.ChildObjects.IndexOf(comment.CommentMarkEnd)

			'Iterate through the child objects of the paragraph within the specified range.
			For i As Integer = startIndex To endIndex

			'Clone each child object.
			Dim doobj As DocumentObject = para.ChildObjects(i).Clone()

			'Add the cloned object to a new paragraph in the destination section.
			destinationSec.AddParagraph().ChildObjects.Add(doobj)
			Next i

			'Save the destination document to a file named "Output.docx" in DOCX format.
			destinationDoc.SaveToFile("Output.docx", FileFormat.Docx)

			'Dispose of the source and destination documents to free up resources.
			sourceDoc.Dispose()
			destinationDoc.Dispose()

			'Launch the Word file.
			WordDocViewer("Output.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
