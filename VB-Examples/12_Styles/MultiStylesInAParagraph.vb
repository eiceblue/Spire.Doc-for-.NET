Imports System.ComponentModel
Imports System.Drawing.Imaging
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace MultiStylesInAParagraph
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a Word document
			Dim doc As New Document()

			'Add a section
			Dim section As Section = doc.AddSection()

			'Add a paragraph
			Dim para As Paragraph = section.AddParagraph()

			'Add a text range 1 and set its style
			Dim range As TextRange = para.AppendText("Spire.Doc for .NET ")
			range.CharacterFormat.FontName = "Calibri"
			range.CharacterFormat.FontSize = 16.0F
			range.CharacterFormat.TextColor = Color.Blue
			range.CharacterFormat.Bold = True
			range.CharacterFormat.UnderlineStyle = UnderlineStyle.Single

			'Add a text range 2 and set its style
			range = para.AppendText("is a professional Word .NET library")
			range.CharacterFormat.FontName = "Calibri"
			range.CharacterFormat.FontSize = 15.0F

			'Save the Word document
			Dim output As String = "MultiStylesInAParagraph_out.docx"
			doc.SaveToFile(output, FileFormat.Docx2013)

			'Dispose of the document object
			doc.Dispose()

			'Launch the file
			FileViewer(output)
		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
