Imports Spire.Doc
Imports Spire.Doc.Interface

Namespace ReplaceWithDocument
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Create a new Document object and load the Word document from the specified file path.
			Dim doc As New Document("..\..\..\..\..\..\Data\Text2.docx")

			' Create a new IDocument object and load another Word document from the specified file path.
			Dim replaceDoc As IDocument = New Document("..\..\..\..\..\..\Data\Text1.docx")

			' Replace the text "Document1" in doc with the contents of replaceDoc.
			' The replacement is not case-sensitive, and multiple occurrences will be replaced.
			doc.Replace("Document1", replaceDoc, False, True)

			' Specify the output file name for saving the modified document.
			Dim output As String = "ReplaceWithDocument.docx"

			' Save the modified document to the specified output file in Docx format.
			doc.SaveToFile(output, FileFormat.Docx)

			' Dispose of the Document object to release resources.
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
