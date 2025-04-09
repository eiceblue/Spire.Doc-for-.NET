Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RepeatRowOnEachPage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Create a new Word document
			Dim document As New Document()

			' Add a section to the document
			Dim section As Section = document.AddSection()

			' Add a table to the section
			Dim table As Table = section.AddTable(True)

			' Set the preferred width of the table to 100%
			Dim width As New PreferredWidth(WidthType.Percentage, 100)
			table.PreferredWidth = width

			' Add a header row to the table
			Dim row As TableRow = table.AddRow()
			row.IsHeader = True
			' Add a cell to the header row
			Dim cell As TableCell = row.AddCell()
			cell.SetCellWidth(100, CellWidthType.Percentage)

			For i As Integer = 0 To row.Cells.Count - 1
				row.Cells(i).CellFormat.Shading.BackgroundPatternColor = Color.LightGray
			Next i

			' Add a paragraph to the cell with text "Row Header 1"
			Dim paragraph As Paragraph = cell.AddParagraph()
			paragraph.AppendText("Row Header 1")
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			' Add another header row to the table
			row = table.AddRow(False, 1)
			row.IsHeader = True
			For i As Integer = 0 To row.Cells.Count - 1
				row.Cells(i).CellFormat.Shading.BackgroundPatternColor = Color.Ivory
			Next i
			row.Height = 30
			cell = row.Cells(0)
			cell.SetCellWidth(100, CellWidthType.Percentage)
			cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle

			' Add a paragraph to the cell with text "Row Header 2"
			paragraph = cell.AddParagraph()
			paragraph.AppendText("Row Header 2")
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			' Add rows and cells to the table
			For i As Integer = 0 To 69
				row = table.AddRow(False, 2)
				cell = row.Cells(0)
				cell.SetCellWidth(50, CellWidthType.Percentage)
				cell.AddParagraph().AppendText("Column 1 Text")
				cell = row.Cells(1)
				cell.SetCellWidth(50, CellWidthType.Percentage)
				cell.AddParagraph().AppendText("Column 2 Text")
			Next i

			' Set background color for alternating rows
			For j As Integer = 1 To table.Rows.Count - 1
				If j Mod 2 = 0 Then
					Dim row2 As TableRow = table.Rows(j)
					For f As Integer = 0 To row2.Cells.Count - 1
						row2.Cells(f).CellFormat.Shading.BackgroundPatternColor = Color.LightBlue
					Next f
				End If
			Next j

			' Save the document to a file
			Dim result As String = "RepeatRowOnEachPage_out.docx"
			document.SaveToFile(result, FileFormat.Docx)

			' Dispose of the document object
			document.Dispose()

			'Launching the Word file.
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
