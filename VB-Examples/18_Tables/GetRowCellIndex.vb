Imports Spire.Doc
Imports System.Text
Imports System.IO

Namespace GetRowCellIndex
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
		   ' Create a new Document object
		   Dim doc As New Document()

		   ' Load an existing Word document from a file
		   doc.LoadFromFile("..\..\..\..\..\..\Data\ReplaceTextInTable.docx")

		   ' Get the first section of the document
		   Dim section As Section = doc.Sections(0)

		   ' Get the first table in the section
		   Dim table As Table = TryCast(section.Tables(0), Table)

		   ' Create a StringBuilder to store the output content
		   Dim content As New StringBuilder()

		   ' Get the collection of tables in the section
		   Dim collections As Spire.Doc.Collections.TableCollection = section.Tables

		   ' Get the index of the table in the collection
		   Dim tableIndex As Integer = collections.IndexOf(table)

		   ' Get the last row in the table and its index
		   Dim row As TableRow = table.LastRow
		   Dim rowIndex As Integer = row.GetRowIndex()

		   ' Get the last cell in the row and its index
		   Dim cell As TableCell = TryCast(row.LastChild, TableCell)
		   Dim cellIndex As Integer = cell.GetCellIndex()

		   ' Append the table, row, and cell indices to the output content
		   content.AppendLine("Table index is " & tableIndex.ToString())
		   content.AppendLine("Row index is " & rowIndex.ToString())
		   content.AppendLine("Cell index is " & cellIndex.ToString())

		   ' Specify the output file path
		   Dim output As String = "GetRowCellIndex_out.txt"

		   ' Write the output content to the output file
		   File.WriteAllText(output, content.ToString())

		   ' Dispose of the document object to free up resources
		   doc.Dispose()

			'Launch the file
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
