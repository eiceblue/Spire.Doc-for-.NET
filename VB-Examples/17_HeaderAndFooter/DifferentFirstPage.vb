Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace DifferentFirstPage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\MultiplePages.docx"

			'Create a word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the first section
			Dim section As Section = doc.Sections(0)

			'specify that the current section has a different header/footer for the first page
			section.PageSetup.DifferentFirstPageHeaderFooter = True

			'Set the first page header. Here we append a picture in the header
			Dim paragraph1 As Paragraph = section.HeadersFooters.FirstPageHeader.AddParagraph()

			'Set horizontal alignment for the paragraph
			paragraph1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right

			'Append a picture
			Dim headerimage As DocPicture = paragraph1.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\E-iceblue.png"))
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim headerimage As DocPicture = paragraph1.AppendPicture("..\..\..\..\..\..\Data\E-iceblue.png")
			' =============================================================================

			'Set the first page footer
			Dim paragraph2 As Paragraph = section.HeadersFooters.FirstPageFooter.AddParagraph()

			'Set horizontal alignment for the paragraph
			paragraph2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			'Append text
			Dim FF As TextRange = paragraph2.AppendText("First Page Footer")

			'Set font size
			FF.CharacterFormat.FontSize = 10

			'Set the other header & footer. If you only need the first page header & footer, don't set this
			Dim paragraph3 As Paragraph = section.HeadersFooters.Header.AddParagraph()

			'Set horizontal alignment for the paragraph
			paragraph3.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			'Append text
			Dim NH As TextRange = paragraph3.AppendText("Spire.Doc for .NET")

			'Set font size
			NH.CharacterFormat.FontSize = 10

			'Add a paragraph
			Dim paragraph4 As Paragraph = section.HeadersFooters.Footer.AddParagraph()

			'Set horizontal alignment for the paragraph
			paragraph4.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			'Append text
			Dim NF As TextRange = paragraph4.AppendText("E-iceblue")

			'Set font size
			NF.CharacterFormat.FontSize = 10

			'Save the document
			Dim output As String = "DifferentFirstPage.docx"
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
