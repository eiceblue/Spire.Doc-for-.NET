Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace CreateTableDirectly
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Create a new Document object
			Dim doc As New Document()

			' Add a new section to the document
			Dim section As Section = doc.AddSection()

			' Create a new table with the document as its parent
			Dim table As New Table(doc)
			table.ResetCells(1, 2)

			' Set the preferred width of the table to 100% of the page width
			table.PreferredWidth = New PreferredWidth(WidthType.Percentage, CShort(100))

			' Set the border type of the table to single line
			table.Format.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single

			' Create a new row for the table
			Dim row As TableRow = table.Rows(0)

			' Set the height of the row to 50.0f
			row.Height = 50.0f

			' Create the first cell of the row
			Dim cell1 As TableCell = table.Rows(0).Cells(0)
			Dim para1 As Paragraph = cell1.AddParagraph()
			' Add text to the cell
			para1.AppendText("Row 1, Cell 1")
			' Set the horizontal alignment of the text
			para1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center
			' Set the background color of the cell
			cell1.CellFormat.Shading.BackgroundPatternColor = Color.CadetBlue
			' Set the vertical alignment of the content in the cell
			cell1.CellFormat.VerticalAlignment = VerticalAlignment.Middle

			' Create the second cell of the row
			Dim cell2 As TableCell = table.Rows(0).Cells(1)
			Dim para2 As Paragraph = cell2.AddParagraph()
			' Add text to the cell
			para2.AppendText("Row 1, Cell 2")
			' Set the horizontal alignment of the text
			para2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center
			' Set the background color of the cell
			cell2.CellFormat.Shading.BackgroundPatternColor = Color.CadetBlue
			' Set the vertical alignment of the content in the cell
			cell2.CellFormat.VerticalAlignment = VerticalAlignment.Middle

			' Add the table to the section
			section.Tables.Add(table)

			' Save the document to a file in Docx2013 format
			Dim output As String = "CreateTableDirectly_out.docx"
			doc.SaveToFile(output, FileFormat.Docx2013)

			' Dispose of the document object to free up resources
			doc.Dispose()

			'Launch the document
			WordDocViewer(output)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
