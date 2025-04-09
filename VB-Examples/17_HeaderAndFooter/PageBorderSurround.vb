Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace PageBorderSurround
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new document
			Dim doc As New Document()

			'Add a section to the document
			Dim section As Section = doc.AddSection()

			'Set the page border properties
			section.PageSetup.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Wave
			section.PageSetup.Borders.Color = Color.Green
			section.PageSetup.Borders.Left.Space = 20
			section.PageSetup.Borders.Right.Space = 20

			'Add a header paragraph to the section
			Dim paragraph1 As Paragraph = section.HeadersFooters.Header.AddParagraph()

			'Set horizontal alignment for the paragraph
			paragraph1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right

			'Append text
			Dim headerText As TextRange = paragraph1.AppendText("Header isn't included in page border")

			'Set the character format for the text
			headerText.CharacterFormat.FontName = "Calibri"
			headerText.CharacterFormat.FontSize = 20
			headerText.CharacterFormat.Bold = True

			'Add a footer paragraph to the section
			Dim paragraph2 As Paragraph = section.HeadersFooters.Footer.AddParagraph()

			'Set horizontal alignment for the paragraph
			paragraph2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left

			'Append text
			Dim footerText As TextRange = paragraph2.AppendText("Footer is included in page border")

			'Set the character format for the text
			footerText.CharacterFormat.FontName = "Calibri"
			footerText.CharacterFormat.FontSize = 20
			footerText.CharacterFormat.Bold = True

			'Set the header not included in the page border while the footer included
			section.PageSetup.PageBorderIncludeHeader = False
			section.PageSetup.HeaderDistance = 40
			section.PageSetup.PageBorderIncludeFooter = True
			section.PageSetup.FooterDistance = 40

			'Save the document
			Dim output As String = "PageBorderSurround.docx"
			doc.SaveToFile(output, FileFormat.Docx)

			'Dispose the document
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
