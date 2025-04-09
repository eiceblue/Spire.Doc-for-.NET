Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Fields

Namespace ReplyToComment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a word document
			Dim doc As New Document()

			'Load the document from disk.
			doc.LoadFromFile("..\..\..\..\..\..\Data\Comment.docx")

			'get the first comment.
			Dim comment1 As Comment = doc.Comments(0)

			'create a new comment and specify the author and content.
			Dim replyComment1 As New Comment(doc)
			replyComment1.Format.Author = "E-iceblue"
			replyComment1.Body.AddParagraph().AppendText("Spire.Doc is a professional Word .NET library on operating Word documents.")

			'add the new comment as a reply to the selected comment.
			comment1.ReplyToComment(replyComment1)

			'Create a DocPicture instance
			Dim docPicture As New DocPicture(doc)

			'Load an image
			Dim img As Image = Image.FromFile("..\..\..\..\..\..\Data\logo.png")
			docPicture.LoadImage(img)

			'insert the picture in the comment
			replyComment1.Body.Paragraphs(0).ChildObjects.Add(docPicture)

			'Save the document.
			doc.SaveToFile("ReplyToComment.docx", FileFormat.Docx)

			'Dispose the document
			doc.Dispose()
			
			'Launch the Word file.
			WordDocViewer("ReplyToComment.docx")
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
