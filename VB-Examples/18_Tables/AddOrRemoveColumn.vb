Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace AddOrRemoveColumn
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			'Create a Word document
			Dim doc As New Document()

			'Load the document from disk
			doc.LoadFromFile("..\..\..\..\..\..\Data\Template_N2.docx")

			'Access the first section
			Dim section As Section = doc.Sections(0)

			'Access the first table
			Dim table As Table = TryCast(section.Tables(0), Table)

			'Add a blank column
			Dim columnIndex1 As Integer = 0
			AddColumn(table, columnIndex1)

			'Remove a column
			Dim columnIndex2 As Integer = 2
			RemoveColumn(table, columnIndex2)

			'Save the file
			Dim output As String = "AddOrRemoveColumn_out.docx"
			doc.SaveToFile(output, FileFormat.Docx2013)

			'Dispose the document
			doc.Dispose()

			'Launch the file
			FileViewer(output)
		End Sub
		Private Sub AddColumn(ByVal table As Table, ByVal columnIndex As Integer)
			For r As Integer = 0 To table.Rows.Count - 1

				'Create a new table cell
				Dim addCell As New TableCell(table.Document)

				'Insert the new cell into the specified position
				table.Rows(r).Cells.Insert(columnIndex, addCell)
			Next r
		End Sub
		Private Sub RemoveColumn(ByVal table As Table, ByVal columnIndex As Integer)
			For r As Integer = 0 To table.Rows.Count - 1

				'Remove the cell from specified position
				table.Rows(r).Cells.RemoveAt(columnIndex)
			Next r
		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
