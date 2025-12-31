Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetVerticalAlignment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			   ' Create a new document object
			   Dim doc As New Document()

			   ' Add a section to the document
			   Dim section As Section = doc.AddSection()

			   ' Add a table to the section with auto-fit behavior
			   Dim table As Table = section.AddTable(True)

			   ' Reset the table cells to 3 rows and 3 columns
			   table.ResetCells(3, 3)

			   ' Apply vertical merging to the first column of the table, spanning 3 rows
			   table.ApplyVerticalMerge(0, 0, 2)

			   ' Set the vertical alignment of cells in the table
			   table.Rows(0).Cells(0).CellFormat.VerticalAlignment = VerticalAlignment.Middle
			   table.Rows(0).Cells(1).CellFormat.VerticalAlignment = VerticalAlignment.Top
			   table.Rows(0).Cells(2).CellFormat.VerticalAlignment = VerticalAlignment.Top
			   table.Rows(1).Cells(1).CellFormat.VerticalAlignment = VerticalAlignment.Middle
			   table.Rows(1).Cells(2).CellFormat.VerticalAlignment = VerticalAlignment.Middle
			   table.Rows(2).Cells(1).CellFormat.VerticalAlignment = VerticalAlignment.Bottom
			   table.Rows(2).Cells(2).CellFormat.VerticalAlignment = VerticalAlignment.Bottom

			   ' Add a paragraph to the first cell of the first row, and append an image to it
			   Dim paraPic As Paragraph = table.Rows(0).Cells(0).AddParagraph()
			   Dim pic As DocPicture = paraPic.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\E-iceblue.png"))
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim pic As DocPicture = paraPic.AppendPicture("..\..\..\..\..\..\Data\E-iceblue.png")
			' =============================================================================

			' Define data for the table cells
			Dim data()() As String = { New String() {"", "Spire.Office", "Spire.DataExport"}, New String() {"", "Spire.Doc", "Spire.DocViewer"}, New String() {"", "Spire.XLS", "Spire.PDF"} }

			   ' Fill the table with data and set cell widths
			   For r As Integer = 0 To 2
				   Dim dataRow As TableRow = table.Rows(r)
				   dataRow.Height = 50
				   For c As Integer = 0 To 2
					   If c = 1 Then
						   ' Add text to the cell and set its width
						   Dim par As Paragraph = dataRow.Cells(c).AddParagraph()
						   par.AppendText(data(r)(c))
						   dataRow.Cells(c).SetCellWidth((section.PageSetup.ClientWidth) / 2, CellWidthType.Point)
					   End If
					   If c = 2 Then
						   ' Add text to the cell and set its width
						   Dim par As Paragraph = dataRow.Cells(c).AddParagraph()
						   par.AppendText(data(r)(c))
						   dataRow.Cells(c).SetCellWidth((section.PageSetup.ClientWidth) / 2, CellWidthType.Point)
					   End If
				   Next c
			   Next r

			   ' Specify the output file name
			   Dim output As String = "SetVerticalAlignment.docx"

			   ' Save the document to a file in Docx format
			   doc.SaveToFile(output, FileFormat.Docx)

			   ' Dispose of the document object
			   doc.Dispose()
	   
			Viewer(output)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
