Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace CombineAndSplitTables
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Combine tables
			CombineTables()

			'Split a table
			SplitTable()
		End Sub
		Private Sub CombineTables()
			'Create a Word document
			Dim doc As New Document()

			'Load document from disk
			doc.LoadFromFile("..\..\..\..\..\..\Data\CombineAndSplitTables.docx")

			'Get the first section
			Dim section As Section = doc.Sections(0)

			'Get the first and second table
			Dim table1 As Table = TryCast(section.Tables(0), Table)
			Dim table2 As Table = TryCast(section.Tables(1), Table)

			'Add the rows of table2 to table1
			For i As Integer = 0 To table2.Rows.Count - 1
				table1.Rows.Add(table2.Rows(i).Clone())
			Next i

			'Remove the table2
			section.Tables.Remove(table2)

			'Save the Word file
			Dim output As String = "CombineTables_out.docx"
			doc.SaveToFile(output, FileFormat.Docx2013)
			
			'Dispose the document
			doc.Dispose()

			'Launch the file
			WordDocViewer(output)

		End Sub
		Private Sub SplitTable()
			'Create a Word document
			Dim doc As New Document()

			'Load document from disk
			doc.LoadFromFile("..\..\..\..\..\..\Data\CombineAndSplitTables.docx")

			'Get the first section
			Dim section As Section = doc.Sections(0)

			'Get the first table
			Dim table As Table = TryCast(section.Tables(0), Table)

			'We will split the table at the third row;
			Dim splitIndex As Integer = 2

			'Create a new table for the split table
			Dim newTable As New Table(section.Document)

			'Add rows to the new table
			For i As Integer = splitIndex To table.Rows.Count - 1
				newTable.Rows.Add(table.Rows(i).Clone())
			Next i

			'Remove rows from original table
			For i As Integer = table.Rows.Count - 1 To splitIndex Step -1
				table.Rows.RemoveAt(i)
			Next i

			'Add the new table in section
			section.Tables.Add(newTable)

			'Save the Word file
			Dim output As String = "SplitTable_out.docx"
			doc.SaveToFile(output, FileFormat.Docx2013)

			'Dispose the document
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
