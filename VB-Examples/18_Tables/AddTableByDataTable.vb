Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AddTableByDataTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a Word document
			Dim document As New Document()

			'Get the first section
			Dim section As Section = document.AddSection()

			'Create a new ParagraphStyle instance
			Dim style As New ParagraphStyle(document)

			'Set the CharacterFormat of the style
			style.CharacterFormat.FontSize = 20f
			style.CharacterFormat.Bold = True
			style.CharacterFormat.TextColor = Color.CadetBlue

			'Add the style to document
			document.Styles.Add(style)

			'Create a paragraph 
			Dim para As Paragraph = section.AddParagraph()

			'Append text
			para.AppendText("Table")

			'Set horizontal alignment for the paragraph
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			'Apply the new style
			para.ApplyStyle(style.Name)

			'Create a DataSet instance
			Dim ds As New DataSet()

			'Load data from a xml file
			ds.ReadXml("..\..\..\..\..\..\Data\dataTable.xml")

			'Get the first data table
			Dim dataTable As DataTable = ds.Tables(0)

			'Add a table
			Dim table As Table = section.AddTable(True)

			'Set its width
			table.PreferredWidth = New PreferredWidth(WidthType.Percentage, 100)

			'Fill table with the data of datatable
			FillTableUsingDataTable(table, dataTable)

			'Set table style
			table.Format.Paddings.All = 5

			For i As Integer = 0 To table.FirstRow.Cells.Count - 1
				table.FirstRow.Cells(i).CellFormat.Shading.BackgroundPatternColor = Color.CadetBlue
			Next i

			'Save the Word file
			Dim output As String = "AddTableUsingDataTable_out.docx"
			document.SaveToFile(output, FileFormat.Docx2013)

			'Dispose the document
			document.Dispose()

			'Launch the file
			FileViewer(output)
		End Sub
		Private Shared Sub FillTableUsingDataTable(ByVal table As Table, ByVal dataTable As DataTable)
			'Get the count of the columns
			Dim columnCount As Integer = dataTable.Columns.Count

			'Loop through the rows of data table
			For Each dataRow As DataRow In dataTable.Rows
				Dim row As TableRow = table.AddRow(columnCount)
				For Each dataColumn As DataColumn In dataTable.Columns

					'Get the column index
					Dim columnIndex As Integer = dataTable.Columns.IndexOf(dataColumn)

					'Get the value 
					Dim value As String = dataRow(dataColumn).ToString()

					'Get the cell object
					Dim cell As TableCell = row.Cells(columnIndex)
					'Add paragraph for cell
					Dim para As Paragraph = cell.AddParagraph()
					'Append text from datatable
					para.AppendText(value)
					'Set the alignment of cell
					cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
					para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center
				Next dataColumn
			Next dataRow
		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
