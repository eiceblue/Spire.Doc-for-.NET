Imports Spire.Doc

Namespace ObtainTableColumnsOfBookmark
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class.
			Dim document As New Document()

			' Load a Word document from the specified relative file path.
			document.LoadFromFile("..\..\..\..\..\..\Data\ObtainTableColumnsOfBookmark.docx")

			' Retrieve the bookmark from the document.
			Dim bookmark As Bookmark = document.Bookmarks.FindByName("tableTest")

			' Get the index of the first column covered by the bookmark.
			Dim firstColumn As Integer = bookmark.FirstColumn

			' Get the index of the last column covered by the bookmark.
			Dim lastColumn As Integer = bookmark.LastColumn

			' Obtain the table cell that contains the start of the bookmark.
			Dim tableCell_start As TableCell = TryCast(bookmark.BookmarkStart.OwnerParagraph.Owner, TableCell)

			' Get the table row that owns the starting cell.
			Dim tableRow_start As TableRow = tableCell_start.OwnerRow

			' Get the table that contains the starting row.
			Dim table As Table = TryCast(tableRow_start.Owner, Table)

			' Determine the row index of the starting row within the table.
			Dim rowIndex_start As Integer = table.Rows.IndexOf(tableRow_start)

			' Obtain the table cell that contains the end of the bookmark.
			Dim tableCell_end As TableCell = TryCast(bookmark.BookmarkEnd.OwnerParagraph.Owner, TableCell)
			' Get the table row that owns the ending cell.
			Dim tableRow_end As TableRow = tableCell_end.OwnerRow
			' Determine the row index of the ending row within the table.
			Dim rowIndex_end As Integer = table.Rows.IndexOf(tableRow_end)

			' Set the text of the starting cell (first column of the start row) to indicate bookmark start.
			table.Rows(rowIndex_start).Cells(firstColumn).FirstParagraph.Text = "Bookmark starts"
			' Highlight the starting cell with a yellow background.
			table.Rows(rowIndex_start).Cells(firstColumn).CellFormat.Shading.BackgroundPatternColor = Color.Yellow
			' Set the text of the ending cell (last column of the end row) to indicate bookmark completion.
			table.Rows(rowIndex_end).Cells(lastColumn).FirstParagraph.Text = "Bookmark completed"
			' Highlight the ending cell with a yellow background.
			table.Rows(rowIndex_end).Cells(lastColumn).CellFormat.Shading.BackgroundPatternColor = Color.Yellow

			' Save the modified document to a new file in .docx format.
			document.SaveToFile("ObtainTableColumnsOfBookmark.docx", FileFormat.Docx)

			' Close the document to release resources.
			document.Close()

			' Explicitly dispose of the document object to free memory.
			document.Dispose()

			'Launching the Word file.
			WordDocViewer("ObtainTableColumnsOfBookmark.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
