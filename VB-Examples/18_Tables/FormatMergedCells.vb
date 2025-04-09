Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace FormatMergedCells
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Add a new section to the document
			Dim section As Section = document.AddSection()

			' Add a table to the section using the AddTable method
			Dim table As Table = AddTable(section)

			' Create a new ParagraphStyle and customize its formatting properties
			Dim style As New ParagraphStyle(document)
			style.Name = "Style"
			style.CharacterFormat.TextColor = Color.DeepSkyBlue
			style.CharacterFormat.Italic = True
			style.CharacterFormat.Bold = True
			style.CharacterFormat.FontSize = 13
			document.Styles.Add(style)

			' Apply horizontal merge for the cells in the first row from column index 0 to 1
			table.ApplyHorizontalMerge(0, 0, 1)

			' Apply the custom style to the paragraph in the first cell of the first row
			table.Rows(0).Cells(0).Paragraphs(0).ApplyStyle(style.Name)

			' Set the vertical alignment and horizontal alignment of the first cell in the first row
			table.Rows(0).Cells(0).CellFormat.VerticalAlignment = VerticalAlignment.Middle
			table.Rows(0).Cells(0).Paragraphs(0).Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			' Apply vertical merge for the cells in the second row from row index 1 to 3
			table.ApplyVerticalMerge(0, 1, 3)

			' Apply the custom style to the paragraph in the first cell of the second row
			table.Rows(1).Cells(0).Paragraphs(0).ApplyStyle(style.Name)

			' Set the vertical alignment and horizontal alignment of the first cell in the second row
			table.Rows(1).Cells(0).CellFormat.VerticalAlignment = VerticalAlignment.Middle
			table.Rows(1).Cells(0).Paragraphs(0).Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left

			' Set the width of the first cell in the second row as a percentage of the table width
			table.Rows(1).Cells(0).SetCellWidth(20, CellWidthType.Percentage)

			' Save the document to a file in Docx format
			Dim output As String = "FormatMergedCells.docx"
			document.SaveToFile(output, FileFormat.Docx)

			' Dispose of the document object to free up resources
			document.Dispose()

			'Launching the file
			WordDocViewer(output)

		End Sub

		Private Shared Function AddTable(ByVal section As Section) As Table
			' Create a new table with 4 rows and 3 columns
			Dim table As Table = section.AddTable(True)
			table.ResetCells(4, 3)

			' Create a DataTable with column headers and data
			Dim dt As New DataTable()
			dt.Columns.Add()
			dt.Columns.Add()
			dt.Columns.Add()
			dt.Rows.Add("Product", "", "Price")
			dt.Rows.Add("Spire.Doc", "Pro Edition", "$799")
			dt.Rows.Add("", "Standard Edition", "$599")
			dt.Rows.Add("", "Free Edition", "$0")

			' Populate the table cells with data from the DataTable
			For r As Integer = 0 To dt.Rows.Count - 1
				Dim dataRow As TableRow = table.Rows(r)
				dataRow.Height = 20
				dataRow.HeightType = TableRowHeightType.Exactly
				For i As Integer = 0 To dataRow.Cells.Count - 1
					dataRow.Cells(i).CellFormat.Shading.BackgroundPatternColor = Color.Empty
				Next i
				For c As Integer = 0 To dataRow.Cells.Count - 1
					If Not String.IsNullOrEmpty(dt.Rows(r)(c).ToString()) Then
						Dim range As TextRange = dataRow.Cells(c).AddParagraph().AppendText(dt.Rows(r)(c).ToString())
						range.CharacterFormat.FontName = "Arial"
					End If
				Next c
			Next r

			Return table
		End Function

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
