Imports Spire.Doc
Imports Spire.Doc.Interface

Namespace DifferentBorders
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click


			' Create a new Document object
			Dim document As New Document()

			' Load an existing Word document from a file
			document.LoadFromFile("..\..\..\..\..\..\Data\TableSample.docx")

			' Get the first table in the document's first section
			Dim table As Table = TryCast(document.Sections(0).Tables(0), Table)

			' Set borders for the entire table
			setTableBorders(table)

			' Set borders for a specific cell in the table
			setCellBorders(table.Rows(2).Cells(0))

			' Save the modified document to a new file
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			' Dispose of the document object to free up resources
			document.Dispose()

			'Launch the MS Word file
			WordDocViewer("Sample.docx")
		End Sub

		Private Sub setTableBorders(ByVal table As Table)
			table.Format.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single
			table.Format.Borders.LineWidth = 3.0F
			table.Format.Borders.Color = Color.Red
		End Sub

		Private Sub setCellBorders(ByVal tableCell As TableCell)
			tableCell.CellFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.DotDash
			tableCell.CellFormat.Borders.LineWidth = 1.0F
			tableCell.CellFormat.Borders.Color = Color.Green
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
