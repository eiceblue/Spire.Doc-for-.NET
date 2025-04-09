Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Collections
Imports System.Text
Imports Spire.Doc.Layout

Namespace SetCommentDisplayMode
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the document from a file
			Dim document As New Document("..\..\..\..\..\..\..\Data\CommentSample.docx")

			' Set comment display mode when converting to pdf
			document.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations
			'document.LayoutOptions.CommentDisplayMode = CommentDisplayMode.Hide;
			'document.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInBalloons;


			document.SaveToFile("SetCommentDisplayMode.pdf",FileFormat.PDF)
			' Dispose the document object
			document.Dispose()

			'Launch result file
			WordDocViewer("SetCommentDisplayMode.pdf")

		End Sub


		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
