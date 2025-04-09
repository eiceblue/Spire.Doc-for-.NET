Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace CellMergeStatus
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim input As String = "..\..\..\..\..\..\Data\CellMergeStatus.docx"

			'Create a Word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the first section
			Dim section As Section = doc.Sections(0)

			'Get the first table in the section
			Dim table As Table = TryCast(section.Tables(0), Table)

			'Create a StringBuilder instance
			Dim stringBuidler As New StringBuilder()

			'Loop through the table rows
			For i As Integer = 0 To table.Rows.Count - 1
				Dim tableRow As TableRow = table.Rows(i)
				For j As Integer = 0 To tableRow.Cells.Count - 1

					'Get each cell
					Dim tableCell As TableCell = tableRow.Cells(j)

					'Returns the way of vertical merging of the cell
					Dim verticalMerge As CellMerge = tableCell.CellFormat.VerticalMerge

					'Get the status of cell merge 
					Dim horizontalMerge As Short = tableCell.GridSpan
					If verticalMerge.Equals(CellMerge.None) AndAlso horizontalMerge = 1 Then
						stringBuidler.Append("Row " & i & ", cell " & j & ": ")
						stringBuidler.AppendLine("This cell isn't merged.")
					Else
						stringBuidler.Append("Row " & i & ", cell " & j & ": ")
						stringBuidler.AppendLine("This cell is merged.")
					End If
				Next j

				'Append an empty line
				stringBuidler.AppendLine()
			Next i

			'Save the document
			Dim output As String = "CellMergeStatus.txt"
			File.WriteAllText(output, stringBuidler.ToString())

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
