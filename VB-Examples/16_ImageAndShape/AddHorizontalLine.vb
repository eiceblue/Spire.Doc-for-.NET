Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields.Shape

Namespace AddHorizontalLine
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a Word document
			Dim doc As New Document()

			'Add a section
			Dim sec As Section = doc.AddSection()

			'Add a paragraph
			Dim para As Paragraph = sec.AddParagraph()

			'Append an horizonal line
			para.AppendHorizonalLine()

			'Save the document
			Dim result As String = "AddHorizontalLine_result.docx"
			doc.SaveToFile(result, FileFormat.Docx)

			'Dispose the document
			doc.Dispose()

			'Launching the MS Word file.
			WordDocViewer(result)
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
