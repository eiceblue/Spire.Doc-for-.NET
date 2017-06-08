Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace InsertBreak
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

            'add cover
            InsertCover(section)

            'insert a break code
            section = document.AddSection()
            section.BreakCode = SectionBreakType.NewPage

            'add content
            InsertContent(section)

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

        Private Sub InsertCover(ByVal section As Section)
            Dim small As ParagraphStyle = New ParagraphStyle(section.Document)
            small.Name = "small"
            small.CharacterFormat.FontName = "Arial"
            small.CharacterFormat.FontSize = 9
            small.CharacterFormat.TextColor = Color.Gray
            section.Document.Styles.Add(small)

            Dim paragraph As Paragraph = section.AddParagraph()
            Paragraph.AppendText("The sample demonstrates how to insert a header and footer into a document.")
            Paragraph.ApplyStyle(small.Name)

            Dim title As Paragraph = section.AddParagraph()
            Dim text As TextRange = title.AppendText("Field Types Supported by Spire.Doc")
            text.CharacterFormat.FontName = "Arial"
            text.CharacterFormat.FontSize = 36
            text.CharacterFormat.Bold = True
            title.Format.BeforeSpacing _
                = section.PageSetup.PageSize.Height / 2 - 3 * section.PageSetup.Margins.Top
            title.Format.AfterSpacing = 8
            title.Format.HorizontalAlignment _
                = Spire.Doc.Documents.HorizontalAlignment.Right

            paragraph = section.AddParagraph()
            paragraph.AppendText("e-iceblue Spire.Doc team.")
            paragraph.ApplyStyle(small.Name)
            paragraph.Format.HorizontalAlignment _
                = Spire.Doc.Documents.HorizontalAlignment.Right
        End Sub

        Private Sub InsertContent(ByVal section As Section)
            Dim list As ParagraphStyle = New ParagraphStyle(section.Document)
            list.Name = "list"
            list.CharacterFormat.FontName = "Arial"
            list.CharacterFormat.FontSize = 11
            list.ParagraphFormat.LineSpacing = 1.5F * 12.0F
            list.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple
            section.Document.Styles.Add(list)

            Dim title As Paragraph = section.AddParagraph()
            Dim text As TextRange = title.AppendText("Field type list:")
            title.ApplyStyle(list.Name)

            Dim first As Boolean = True
            For Each type As FieldType In System.Enum.GetValues(GetType(FieldType))
                If type = FieldType.FieldUnknown _
                    Or type = FieldType.FieldNone Or type = FieldType.FieldEmpty Then
                    Continue For
                End If

                Dim paragraph As Paragraph = section.AddParagraph()
                paragraph.AppendText(String.Format("{0} is supported in Spire.Doc", type))

                If (first) Then
                    paragraph.ListFormat.ApplyNumberedStyle()
                    first = False
                Else
                    paragraph.ListFormat.ContinueListNumbering()
                End If

                paragraph.ApplyStyle(list.Name)
            Next
        End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
