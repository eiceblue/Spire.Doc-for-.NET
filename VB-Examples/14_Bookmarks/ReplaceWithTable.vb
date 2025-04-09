Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ReplaceWithTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\Bookmark.docx"

			'Create a word document
			Dim doc As New Document()

			'Load the document
			doc.LoadFromFile(input)

			'Create a table
			Dim table As New Table(doc, True)

			'Create datatable
			Dim dt As New DataTable()
			dt.Columns.Add("id", GetType(String))
			dt.Columns.Add("name", GetType(String))
			dt.Columns.Add("job", GetType(String))
			dt.Columns.Add("email", GetType(String))
			dt.Columns.Add("salary", GetType(String))
			dt.Rows.Add(New String() {"Name", "Capital", "Continent", "Area", "Population"})
			dt.Rows.Add(New String() {"Argentina", "Buenos Aires", "South America", "2777815", "32300003"})
			dt.Rows.Add(New String() {"Bolivia", "La Paz", "South America", "1098575", "7300000"})
			dt.Rows.Add(New String() {"Brazil", "Brasilia", "South America", "8511196", "150400000"})
			table.ResetCells(dt.Rows.Count, dt.Columns.Count)

			'Fill the table with the data of datatable
			For i As Integer = 0 To dt.Rows.Count - 1
				For j As Integer = 0 To dt.Columns.Count - 1
					table.Rows(i).Cells(j).AddParagraph().AppendText(dt.Rows(i)(j).ToString())
				Next j
			Next i

			'Create a BookmarksNavigator instance
			Dim navigator As New BookmarksNavigator(doc)

			'Get the specific bookmark by its name
			navigator.MoveToBookmark("Test")

			'Create a TextBodyPart instance and add the table to it
			Dim part As New TextBodyPart(doc)

			'Add the table to the textpart
			part.BodyItems.Add(table)

			'Replace the current bookmark content with the TextBodyPart object
			navigator.ReplaceBookmarkContent(part)

			'Save and launch document
			Dim output As String = "ReplaceWithTable.docx"
			doc.SaveToFile(output, FileFormat.Docx)

			'Dispose the document
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
