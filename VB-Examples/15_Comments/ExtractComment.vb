Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ExtractComment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			
			Dim input As String = "..\..\..\..\..\..\Data\CommentSample.docx"

			'Create a Word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Create a StringBuilder instance
			Dim SB As New StringBuilder()

			'Traverse all comments
			For Each comment As Comment In doc.Comments
				For Each p As Paragraph In comment.Body.Paragraphs
					'Append the comments to the StringBuilder instance
					SB.AppendLine(p.Text)
				Next p
			Next comment

			'Save to TXT File
			Dim output As String = "ExtractComment.txt"
			File.WriteAllText(output, SB.ToString())

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
