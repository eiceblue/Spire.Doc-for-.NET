Imports Spire.Doc


Namespace RemoveAndReplaceComment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\CommentSample.docx"

			'Create a word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Replace the content of the first paragraph
			doc.Comments(0).Body.Paragraphs(0).Replace("This is the title", "This comment is changed.", False, False)

			'Remove the second comment
			doc.Comments.RemoveAt(1)

			'Save the document
			Dim output As String = "RemoveAndReplaceComment.docx"
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
