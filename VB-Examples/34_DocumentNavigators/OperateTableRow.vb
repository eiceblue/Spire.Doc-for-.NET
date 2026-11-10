Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace OperateTableRow
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new empty document instance.
			Dim doc As New Document()

			' Create a document navigator to help navigate and modify the document content.
			Dim navigator As New DocumentNavigator(doc)

			' Start creating a new table and get its reference.
			Dim table As Table = navigator.StartTable()

			' Initialize the table with 2 rows and 2 columns.
			table.ResetCells(2, 2)

			' Set the table width to 100% of the available page width.
			table.PreferredWidth = New PreferredWidth(WidthType.Percentage, CShort(100))

			' Apply a single black border with line width 1 to all sides of the table.
			table.Format.Borders.SetBordersAttribute(Spire.Doc.Documents.BorderStyle.Single, 1, Color.Black)

			' Set the height of the first row to 30 points.
			table.FirstRow.Height = 30f

			' Get a reference to the first cell in the first row (row 0, column 0).
			Dim cell1 As TableCell = table.Rows(0).Cells(0)

			' Add content to the first cell of the first row.
			addcellcontent(cell1, "Row 1, Cell 1")

			' Get a reference to the second cell in the first row (row 0, column 1).
			Dim cell2 As TableCell = table.Rows(0).Cells(1)

			' Add content to the second cell of the first row.
			addcellcontent(cell2, "Row 1, Cell 2")

			' Get a reference to the first cell in the second row (row 1, column 0).
			Dim cell3 As TableCell = table.Rows(1).Cells(0)

			' Add content to the first cell of the second row.
			addcellcontent(cell3, "Row 2, Cell 1")

			' Get a reference to the second cell in the second row (row 1, column 1).
			Dim cell4 As TableCell = table.Rows(1).Cells(1)

			' Add content to the second cell of the second row.
			addcellcontent(cell4, "Row 2, Cell 2")

			' Finalize the table creation.
			navigator.EndTable()

			' Delete the first row (row index 0) of the first table (table index 0) in the document.
			navigator.DeleteRow(0, 0)

			' Save the modified document to a new file named "OperateTableRow.docx" in DOCX format.
			doc.SaveToFile("OperateTableRow.docx", FileFormat.Docx)

			' Close the document to release internal resources.
			doc.Close()

			' Explicitly dispose of the document object to free memory immediately.
			doc.Dispose()

			'Launch the MS Word file.
			WordDocViewer("OperateTableRow.docx")
		End Sub

		Private Sub addcellcontent(ByVal cell As TableCell, ByVal Content As String)
			' Add a new paragraph to the table cell.
			Dim para As Paragraph = cell.AddParagraph()

			' Append the specified text content to the paragraph inside the cell.
			para.AppendText(Content)

			' Center-align the text horizontally within the paragraph.
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			' Set the background (shading) color of the table cell to CadetBlue.
			cell.CellFormat.Shading.BackgroundPatternColor = Color.CadetBlue

			' Vertically center the content within the table cell.
			cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
