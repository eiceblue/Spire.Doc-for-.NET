Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetAuthorForRevisions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\..\Data\ExtractText.docx")

			'Get the first section
			Dim section As Section=document.Sections(0)

			' Start track revisions
			document.StartTrackRevisions("test")

			' Set author for deleted revision
			Dim para As Paragraph =document.LastParagraph
			para.Text = ""
			For i As Integer = 0 To para.ChildObjects.Count - 1
				Dim textRange As TextRange = TryCast(para.ChildObjects(i), TextRange)
				If textRange.IsDeleteRevision Then
					textRange.DeleteRevision.Author = "user1"
				End If
			Next i

			' Set author for inserted revision
			Dim paragraph As Paragraph = section.AddParagraph()
			Dim range As TextRange = paragraph.AppendText("Added text")
			range.InsertRevision.Author = "user2"

			' Stop track revisions
			document.StopTrackRevisions()

			' Save the file
			document.SaveToFile("SetAuthorForRevisions_out.docx", FileFormat.Docx)

			' Dispose of the Document object 
			document.Dispose()

		
			WordDocViewer("SetAuthorForRevisions_out.docx")

		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
