Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace StartTrackRevisions
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

			' Start track revisions
			document.StartTrackRevisions("User01", Date.Now)

			' Get the first paragraph and add content
			document.Sections(0).Paragraphs(0).AppendText("User01 add new Text!")

			' Delete a paragraph
			document.Sections(0).Paragraphs.RemoveAt(2)

			' Stop track revisions
			document.StopTrackRevisions()

			' Save the file
			document.SaveToFile("StartTrackRevisions_out.docx", FileFormat.Docx)

			' Dispose of the Document object 
			document.Dispose()

		
			WordDocViewer("StartTrackRevisions_out.docx")

		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
