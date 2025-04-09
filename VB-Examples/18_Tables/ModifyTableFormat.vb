Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ModifyTableFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Create a new Document object
			Dim document As New Document()

			' Load an existing Word document from a file
			document.LoadFromFile("..\..\..\..\..\..\Data\ModifyTableFormat.docx")

			' Get the first section of the document
			Dim section As Section = document.Sections(0)

			' Get the tables in the section
			Dim tb1 As Table = TryCast(section.Tables(0), Table)
			Dim tb2 As Table = TryCast(section.Tables(1), Table)
			Dim tb3 As Table = TryCast(section.Tables(2), Table)

			' Modify the format of tb1
			MoidfyTableFormat(tb1)

			' Modify the row format of tb2
			ModifyRowFormat(tb2)

			' Modify the cell format of tb3
			ModifyCellFormat(tb3)

			' Specify the output file path
			Dim output As String = "ModifyTableFormat_out.docx"

			' Save the modified document to a file
			document.SaveToFile(output, FileFormat.Docx2013)

			' Dispose of the document object to free up resources
			document.Dispose()

			'Launch Word file.
			WordDocViewer(output)
		End Sub
		' Modify the table format
		Private Shared Sub MoidfyTableFormat(ByVal table As Table)
			' Set the preferred width of the table
			table.PreferredWidth = New PreferredWidth(WidthType.Twip, CShort(6000))

			' Apply a specific table style to the table
			table.ApplyStyle(DefaultTableStyle.ColorfulGridAccent3)

			' Set padding for all cells in the table
			table.Format.Paddings.All = 5

			' Set the title and description of the table
			table.Title = "Spire.Doc for .NET"
			table.TableDescription = "Spire.Doc for .NET is a professional Word .NET library"
		End Sub

		' Modify the row format
		Private Shared Sub ModifyRowFormat(ByVal table As Table)
			' Set the cell spacing of the first row
			table.Format.CellSpacing = 2

			' Set the height of the second row
			table.Rows(1).HeightType = TableRowHeightType.Exactly
			table.Rows(1).Height = 20f

			' Set the background color of the third row
			For i As Integer = 0 To table.Rows(2).Cells.Count - 1
				table.Rows(2).Cells(i).CellFormat.Shading.BackgroundPatternColor = Color.DarkSeaGreen
			Next i
		End Sub

		' Modify the cell format
		Private Shared Sub ModifyCellFormat(ByVal table As Table)
			' Set the vertical alignment and horizontal alignment of the first cell in the first row
			table.Rows(0).Cells(0).CellFormat.VerticalAlignment = VerticalAlignment.Middle
			table.Rows(0).Cells(0).Paragraphs(0).Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			' Set the background color of the first cell in the second row
			table.Rows(1).Cells(0).CellFormat.Shading.BackgroundPatternColor = Color.DarkSeaGreen

			' Set borders for the first cell in the third row
			table.Rows(2).Cells(0).CellFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single
			table.Rows(2).Cells(0).CellFormat.Borders.LineWidth = 1f
			table.Rows(2).Cells(0).CellFormat.Borders.Left.Color = Color.Red
			table.Rows(2).Cells(0).CellFormat.Borders.Right.Color = Color.Red
			table.Rows(2).Cells(0).CellFormat.Borders.Top.Color = Color.Red
			table.Rows(2).Cells(0).CellFormat.Borders.Bottom.Color = Color.Red

			' Set the text direction of the first cell in the fourth row
			table.Rows(3).Cells(0).CellFormat.TextDirection = TextDirection.RightToLeft
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace