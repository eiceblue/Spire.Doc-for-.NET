Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace PageSetup
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
            Dim document As New Document()
            Dim section As Section = document.AddSection()

            'the unit of all measures below is point, 1point = 0.3528 mm
            section.PageSetup.PageSize = PageSize.A4
            section.PageSetup.Margins.Top = 72.0F
            section.PageSetup.Margins.Bottom = 72.0F
            section.PageSetup.Margins.Left = 89.85F
            section.PageSetup.Margins.Right = 89.85F

            'insert header and footer
            InsertHeaderAndFooter(section)

            addTable(section)

			'Save doc file.
            document.SaveToFile("Sample.doc", FileFormat.Doc)

			'Launching the MS Word file.
			WordDocViewer("Sample.doc")
        End Sub

        Private Sub addTable(ByVal section As Section)
            Dim header As String() = {"Name", "Capital", "Continent", "Area", "Population"}
            Dim data As String()() = { _
                    New String() {"Argentina", "Buenos Aires", "South America", "2777815", "32300003"}, _
                    New String() {"Bolivia", "La Paz", "South America", "1098575", "7300000"}, _
                    New String() {"Brazil", "Brasilia", "South America", "8511196", "150400000"}, _
                    New String() {"Canada", "Ottawa", "North America", "9976147", "26500000"}, _
                    New String() {"Chile", "Santiago", "South America", "756943", "13200000"}, _
                    New String() {"Colombia", "Bagota", "South America", "1138907", "33000000"}, _
                    New String() {"Cuba", "Havana", "North America", "114524", "10600000"}, _
                    New String() {"Ecuador", "Quito", "South America", "455502", "10600000"}, _
                    New String() {"El Salvador", "San Salvador", "North America", "20865", "5300000"}, _
                    New String() {"Guyana", "Georgetown", "South America", "214969", "800000"}, _
                    New String() {"Jamaica", "Kingston", "North America", "11424", "2500000"}, _
                    New String() {"Mexico", "Mexico City", "North America", "1967180", "88600000"}, _
                    New String() {"Nicaragua", "Managua", "North America", "139000", "3900000"}, _
                    New String() {"Paraguay", "Asuncion", "South America", "406576", "4660000"}, _
                    New String() {"Peru", "Lima", "South America", "1285215", "21600000"}, _
                    New String() {"United States of America", "Washington", "North America", "9363130", "249200000"}, _
                    New String() {"Uruguay", "Montevideo", "South America", "176140", "3002000"}, _
                    New String() {"Venezuela", "Caracas", "South America", "912047", "19700000"} _
                }
            Dim table As Spire.Doc.Table = section.AddTable()
            table.ResetCells(data.Length + 1, header.Length)

            ' ***************** First Row *************************
            Dim row As TableRow = table.Rows(0)
            row.IsHeader = True
            row.Height = 20    'unit: point, 1point = 0.3528 mm
            row.HeightType = TableRowHeightType.Exactly
            row.RowFormat.BackColor = Color.Gray
            For i As Integer = 0 To header.Length - 1
                row.Cells(i).CellFormat.VerticalAlignment = VerticalAlignment.Middle
                Dim p As Paragraph = row.Cells(i).AddParagraph()
                p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center
                Dim txtRange As TextRange = p.AppendText(header(i))
                txtRange.CharacterFormat.Bold = True
            Next


            For r As Integer = 0 To data.Length - 1
                Dim dataRow As TableRow = table.Rows(r + 1)
                dataRow.Height = 20
                dataRow.HeightType = TableRowHeightType.Exactly
                dataRow.RowFormat.BackColor = Color.Empty
                For c As Integer = 0 To data(r).Length - 1
                    dataRow.Cells(c).CellFormat.VerticalAlignment = VerticalAlignment.Middle
                    dataRow.Cells(c).AddParagraph().AppendText(data(r)(c))
                Next
            Next
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

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
