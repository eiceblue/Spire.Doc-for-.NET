Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace InsertPictureIntoComment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\CommentTemplate.docx"

			'Create a word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the third paragraph
			Dim paragraph As Paragraph = doc.Sections(0).Paragraphs(2)

			'Add comment
			Dim comment As Comment = paragraph.AppendComment("This is a comment.")

			'Add author information
			comment.Format.Author = "E-iceblue"

			'Create a DocPicture instance
			Dim docPicture As New DocPicture(doc)

			'Load an Image
			Dim img As Image = Image.FromFile("..\..\..\..\..\..\Data\E-iceblue.png")
			docPicture.LoadImage(img)

			'Insert the picture into the comment body
			comment.Body.AddParagraph().ChildObjects.Add(docPicture)

			'Save and launch
			Dim output As String = "InsertPictureIntoComment.docx"
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
