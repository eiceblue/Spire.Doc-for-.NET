Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Formatting

Namespace OperateTableCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document instance.
			Dim doc As New Document()

			' Initialize a document navigator to navigate and modify the document content.
			Dim navigator As New DocumentNavigator(doc)

			' Begin creating a new table in the document.
			navigator.StartTable()

			' Insert the first cell into the current row of the table and get its reference.
			Dim cell1 As TableCell = navigator.InsertCell()

			' Add content to the first cell of the first row.
			addcellcontent(cell1, "Row 1, Cell 1")

			' Insert the second cell into the current row of the table and get its reference.
			Dim cell2 As TableCell = navigator.InsertCell()

			' Add content to the second cell of the first row.
			addcellcontent(cell2, "Row 1, Cell 2")

			' Insert the third cell into the current row of the table and get its reference.
			Dim cell3 As TableCell = navigator.InsertCell()

			' Add content to the third cell of the first row.
			addcellcontent(cell3, "Row 1, Cell 3")

			' End the current row and move to the next row in the table.
			navigator.EndRow()

			' Insert the first cell into the new row of the table and get its reference.
			Dim cell4 As TableCell = navigator.InsertCell()

			' Add content to the first cell of the second row.
			addcellcontent(cell4, "Row 2, Cell 1")

			' End the table creation process.
			navigator.EndTable()

			' Move the navigator's cursor to a specific cell in the first table
			navigator.MoveToCell(0, 0, 1, 0)

			' Insert (overwrite) the text at the current cursor position inside the target cell.
			navigator.Write("new content")

			' Get the formatting object of the current cell where the navigator is positioned.
			Dim cellformat As CellFormat = navigator.CellFormat

			' Clear all existing formatting applied to the current cell.
			cellformat.ClearFormatting()

			' Set the background (shading) color of the cell to GreenYellow.
			cellformat.Shading.BackgroundPatternColor = Color.GreenYellow

			' Save the modified document to the specified output file in DOCX format.
			doc.SaveToFile("OperateTableCell.docx", FileFormat.Docx)

			' Close the document to release internal resources.
			doc.Close()

			' Explicitly dispose of the document object to free memory immediately.
			doc.Dispose()

			'Launch the MS Word file.
			WordDocViewer("OperateTableCell.docx")
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
