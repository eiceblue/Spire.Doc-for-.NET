Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace HeaderAndFooter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a Word document
			Dim document As New Document()

			'Load the file from disk
			document.LoadFromFile("..\..\..\..\..\..\Data\Sample.docx")

			'Get the first section
			Dim section As Section = document.Sections(0)

			'Insert header and footer
			InsertHeaderAndFooter(section)

			'Save the file.
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			'Dispose the document
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("Sample.docx")
		End Sub

		Private Sub InsertHeaderAndFooter(ByVal section As Section)
			'Get the header
			Dim header As HeaderFooter = section.HeadersFooters.Header

			'Get the footer
			Dim footer As HeaderFooter = section.HeadersFooters.Footer

			'Create a new paragraph for the header and add an image
			Dim headerParagraph As Paragraph = header.AddParagraph()
			Dim headerPicture As DocPicture = headerParagraph.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Header.png"))

			'Add text to the header paragraph and set its formatting properties
			Dim text As TextRange = headerParagraph.AppendText("Demo of Spire.Doc")
			text.CharacterFormat.FontName = "Arial"
			text.CharacterFormat.FontSize = 10
			text.CharacterFormat.Italic = True
			headerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right

			'Set border properties for the bottom border of the header paragraph
			headerParagraph.Format.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single
			headerParagraph.Format.Borders.Bottom.Space = 0.05F


			'Set the text wrapping style and alignment properties for the header picture
			headerPicture.TextWrappingStyle = TextWrappingStyle.Behind
			headerPicture.HorizontalOrigin = HorizontalOrigin.Page
			headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
			headerPicture.VerticalOrigin = VerticalOrigin.Page
			headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top

			'Create a new paragraph for the footer and add an image
			Dim footerParagraph As Paragraph = footer.AddParagraph()
			Dim footerPicture As DocPicture = footerParagraph.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Footer.png"))

			'Set the text wrapping style and alignment properties for the footer picture
			footerPicture.TextWrappingStyle = TextWrappingStyle.Behind
			footerPicture.HorizontalOrigin = HorizontalOrigin.Page
			footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
			footerPicture.VerticalOrigin = VerticalOrigin.Page
			footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom

			'Add fields for page number and total number of pages to the footer paragraph
			footerParagraph.AppendField("page number", FieldType.FieldPage)
			footerParagraph.AppendText(" of ")
			footerParagraph.AppendField("number of pages", FieldType.FieldNumPages)
			footerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right

			'Set border properties for the top border of the footer paragraph
			footerParagraph.Format.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single
			footerParagraph.Format.Borders.Top.Space = 0.05F
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
