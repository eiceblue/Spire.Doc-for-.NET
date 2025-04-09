Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace AddOrDeleteRow
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a document
			Dim document As New Document()

			'Load the file from disk
			document.LoadFromFile("..\..\..\..\..\..\Data\TableSample.docx")

			'Get the first section
			Dim section As Section = document.Sections(0)

			'Get the first table
			Dim table As Table = TryCast(section.Tables(0), Table)

			'Delete the eighth row
			table.Rows.RemoveAt(7)

			'Add a row and insert it into specific position
			Dim row As New TableRow(document)
			For i As Integer = 0 To table.Rows(0).Cells.Count - 1

				'Add a cell
				Dim tc As TableCell = row.AddCell()

				'Add a paragraph for the cell
				Dim paragraph As Paragraph = tc.AddParagraph()

				'Set horizontal alignment for the paragraph
				paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center
				paragraph.AppendText("Added")
			Next i

			'Insert the new row
			table.Rows.Insert(2, row)

			'Add a row at the end of table
			table.AddRow()

			'Save to file
			document.SaveToFile("AddDeleteRow.docx", FileFormat.Docx)

			'Dispose the document
			document.Dispose()
			FileViewer("AddDeleteRow.docx")
		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
