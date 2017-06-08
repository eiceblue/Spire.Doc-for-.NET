Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace Encrypt
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            'Create word document
            Dim document As New Document()

            Dim section As Section = document.AddSection()

            'page setup
            SetPage(section)

            'insert header and footer
            InsertHeaderAndFooter(section)

            'add content
            InsertContent(section)

            'encrypt document with password specified by textBox1
            document.Encrypt(Me.textBox1.Text)

			'Save doc file.
            document.SaveToFile("Sample.doc", FileFormat.Doc)

			'Launching the MS Word file.
			WordDocViewer("Sample.doc")


        End Sub

        Private Sub InsertHeaderAndFooter(ByVal section As Section)
            Dim header As HeaderFooter = section.HeadersFooters.Header
            Dim footer As HeaderFooter = section.HeadersFooters.Footer

            'insert picture and text to header
            Dim headerParagraph As Paragraph = header.AddParagraph()
            Dim headerPicture As DocPicture _
                = headerParagraph.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Header.png"))

            'header text
            Dim text As TextRange = headerParagraph.AppendText("Demo of Spire.Doc")
            Text.CharacterFormat.FontName = "Arial"
            Text.CharacterFormat.FontSize = 10
            Text.CharacterFormat.Italic = True
            headerParagraph.Format.HorizontalAlignment _
                = Spire.Doc.Documents.HorizontalAlignment.Right

            'border
            headerParagraph.Format.Borders.Bottom.BorderType _
                = Spire.Doc.Documents.BorderStyle.Single
            headerParagraph.Format.Borders.Bottom.Space = 0.05F


            'header picture layout - text wrapping
            headerPicture.TextWrappingStyle = TextWrappingStyle.Behind

            'header picture layout - position
            headerPicture.HorizontalOrigin = HorizontalOrigin.Page
            headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
            headerPicture.VerticalOrigin = VerticalOrigin.Page
            headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top

            'insert picture to footer
            Dim footerParagraph As Paragraph = footer.AddParagraph()
            Dim footerPicture As DocPicture _
                = footerParagraph.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Footer.png"))

            'footer picture layout
            footerPicture.TextWrappingStyle = TextWrappingStyle.Behind
            footerPicture.HorizontalOrigin = HorizontalOrigin.Page
            footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
            footerPicture.VerticalOrigin = VerticalOrigin.Page
            footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom

            'insert page number
            footerParagraph.AppendField("page number", FieldType.FieldPage)
            footerParagraph.AppendText(" of ")
            footerParagraph.AppendField("number of pages", FieldType.FieldNumPages)
            footerParagraph.Format.HorizontalAlignment _
                = Spire.Doc.Documents.HorizontalAlignment.Right

            'border
            footerParagraph.Format.Borders.Top.BorderType _
                = Spire.Doc.Documents.BorderStyle.Single
            footerParagraph.Format.Borders.Top.Space = 0.05F
        End Sub

        Private Sub SetPage(ByVal section As Section)
            'the unit of all measures below is point, 1point = 0.3528 mm
            section.PageSetup.PageSize = PageSize.A4
            section.PageSetup.Margins.Top = 72.0F
            section.PageSetup.Margins.Bottom = 72.0F
            section.PageSetup.Margins.Left = 89.85F
            section.PageSetup.Margins.Right = 89.85F
        End Sub

        Private Sub InsertContent(ByVal section As Section)
            'title
            Dim paragraph As Paragraph = section.AddParagraph()
            Dim title As TextRange = paragraph.AppendText("Summary of Science")
            title.CharacterFormat.Bold = True
            title.CharacterFormat.FontName = "Arial"
            title.CharacterFormat.FontSize = 14
            paragraph.Format.HorizontalAlignment _
                = Spire.Doc.Documents.HorizontalAlignment.Center
            paragraph.Format.AfterSpacing = 10

            'style
            Dim style1 As ParagraphStyle = New ParagraphStyle(section.Document)
            style1.Name = "style1"
            style1.CharacterFormat.FontName = "Arial"
            style1.CharacterFormat.FontSize = 9
            style1.ParagraphFormat.LineSpacing = 1.5F * 12.0F
            style1.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple
            section.Document.Styles.Add(style1)

            Dim style2 As ParagraphStyle = New ParagraphStyle(section.Document)
            style2.Name = "style2"
            style2.ApplyBaseStyle(style1.Name)
            style2.CharacterFormat.Font = New Font("Arial", 10.0F)
            section.Document.Styles.Add(style2)

            paragraph = section.AddParagraph()
            paragraph.AppendText("(All text and pictures are from ")
            Dim link As String = "http://en.wikipedia.org/wiki/Science"
            paragraph.AppendHyperlink(link, "Wikipedia", HyperlinkType.WebLink)
            paragraph.AppendText(", the free encyclopedia)")
            paragraph.ApplyStyle(style1.Name)

            Dim paragraph1 As Paragraph = section.AddParagraph()
            Dim str1 As String _
                = "Science (from the Latin scientia, meaning ""knowledge"") " _
                + "is an enterprise that builds and organizes knowledge in the form " _
                + "of testable explanations and predictions about the natural world. " _
                + "An older meaning still in use today is that of Aristotle, " _
                + "for whom scientific knowledge was a body of reliable knowledge " _
                + "that can be logically and rationally explained " _
                + "(see ""History and etymology"" section below)."
            paragraph1.AppendText(str1)

            'Insert a picture in the right of the paragraph1
            Dim picture As DocPicture _
                = paragraph1.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Wikipedia_Science.png"))
            picture.TextWrappingStyle = TextWrappingStyle.Square
            picture.TextWrappingType = TextWrappingType.Left
            picture.VerticalOrigin = VerticalOrigin.Paragraph
            picture.VerticalPosition = 0
            picture.HorizontalOrigin = HorizontalOrigin.Column
            picture.HorizontalAlignment = ShapeHorizontalAlignment.Right

            paragraph1.ApplyStyle(style2.Name)

            Dim paragraph2 As Paragraph = section.AddParagraph()
            Dim str2 As String _
                = "Since classical antiquity science as a type of knowledge was closely linked " _
                + "to philosophy, the way of life dedicated to discovering such knowledge. " _
                + "And into early modern times the two words, ""science"" and ""philosophy"", " _
                + "were sometimes used interchangeably in the English language. " _
                + "By the 17th century, ""natural philosophy"" " _
                + "(which is today called ""natural science"") could be considered separately " _
                + "from ""philosophy"" in general. But ""science"" continued to also be used " _
                + "in a broad sense denoting reliable knowledge about a topic, in the same way " _
                + "it is still used in modern terms such as library science or political science."
            paragraph2.AppendText(str2)
            paragraph2.ApplyStyle(style2.Name)

            Dim paragraph3 As Paragraph = section.AddParagraph()
            Dim str3 As String _
                = "The more narrow sense of ""science"" that is common today developed as a part " _
                + "of science became a distinct enterprise of defining ""laws of nature"", " _
                + "based on early examples such as Kepler's laws, Galileo's laws, and Newton's " _
                + "laws of motion. In this period it became more common to refer to natural " _
                + "philosophy as  ""natural science"". Over the course of the 19th century, the word " _
                + """science"" became increasingly associated with the disciplined study of the " _
                + "natural world including physics, chemistry, geology and biology. This sometimes " _
                + "left the study of human thought and society in a linguistic limbo, which was " _
                + "resolved by classifying these areas of academic study as social science. " _
                + "Similarly, several other major areas of disciplined study and knowledge " _
                + "exist today under the general rubric of ""science"", such as formal science " _
                + "and applied science."
            paragraph3.AppendText(str3)
            paragraph3.ApplyStyle(style2.Name)
        End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
