Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetSpacing
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load the Word document from the specified file path
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_1.docx")

			' Create a new paragraph with the document as its parent
			Dim para As New Paragraph(document)

			' Append text to the paragraph and apply formatting
			Dim textRange1 As TextRange = para.AppendText("This is an inserted paragraph.")
			textRange1.CharacterFormat.TextColor = Color.Blue
			textRange1.CharacterFormat.FontSize = 15

			' Configure spacing settings for the paragraph
			para.Format.BeforeAutoSpacing = False
			para.Format.BeforeSpacing = 10
			para.Format.AfterAutoSpacing = False
			para.Format.AfterSpacing = 10

			' Insert the paragraph at index 1 in the first section's collection of paragraphs
			document.Sections(0).Paragraphs.Insert(1, para)

			' Specify the output file name for the modified document
			Dim result As String = "Result-SetTheSpacing.docx"

			' Save the modified document to the specified file format (Docx2013)
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document object to release resources
			document.Dispose()
			
			'Launch the file.
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
