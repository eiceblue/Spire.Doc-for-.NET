Imports Spire.Doc

Namespace WordToMarkDown
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Define the input file path relative to the current directory.
			Dim input As String = "..\..\..\..\..\..\Data\ToMD.docx"

			' Create a new document
			Dim doc As New Document()

			' Load .docx file
			doc.LoadFromFile(input)

			' Save to .md file
			doc.SaveToFile("WordToMarkDown_output.md", FileFormat.Markdown)

			' Dispose of the Document object
			doc.Close()

		
			WordDocViewer("WordToMarkDown_output.md")

		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
