Imports Spire.Doc

Namespace MarkDownToWordOrPDF
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Define the input file path relative to the current directory.
			Dim input As String = "..\..\..\..\..\..\Data\MarkDownFile.md"

			' Create a new Document object
			Dim doc As New Document()

			'Load .md file
			doc.LoadFromFile(input)

			'Save to .md file
			'doc.SaveToFile("output.md", Spire.Doc.FileFormat.Markdown);
			'Save to .docx file
			'doc.SaveToFile("output.docx", Spire.Doc.FileFormat.Docx);
			'Save to .doc file
			'doc.SaveToFile("output.doc", Spire.Doc.FileFormat.Doc);
			'Save to .pdf file
			doc.SaveToFile("output.pdf",FileFormat.PDF)

			' Dispose of the Document object
			doc.Close()

	
			Viewer("output.pdf")
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
