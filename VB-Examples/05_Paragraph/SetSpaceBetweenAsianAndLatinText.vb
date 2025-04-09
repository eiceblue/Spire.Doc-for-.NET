Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Fields.Shape

Namespace SetSpaceBetweenAsianAndLatinText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load the Word document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\SetSpaceBetweenAsianAndLatinText.docx")

			' Retrieve the first paragraph in the first section of the document
			Dim para As Paragraph = document.Sections(0).Paragraphs(0)

			' Disable automatic spacing between East Asian and Latin characters
			para.Format.AutoSpaceDE = False

			' Enable automatic spacing between East Asian and non-East Asian characters
			para.Format.AutoSpaceDN = True

			' Specify the output file name for the modified document
			Dim result As String = "Result.docx"

			' Save the modified document to the specified file format (Docx2013)
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document object to release resources
			document.Dispose()

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
