Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace AddGutter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load the specified document file
			document.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			' Access the first section of the document
			Dim section As Section = document.Sections(0)

			' Set the gutter size for the page setup of the section
			section.PageSetup.Gutter = 100.0F

			' Specify the file path for the output result
			Dim output As String = "InsertGutter.docx"

			' Save the modified document to a new file with the specified file format
			document.SaveToFile(output, FileFormat.Docx)

			' Dispose of the document object
			document.Dispose()

			'Launching the file
			WordDocViewer(output)

		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
