Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.Text
Imports System.IO

Namespace GetDiagonalBorder
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load an existing Word document from a file
			document.LoadFromFile("..\..\..\..\..\..\Data\GetDiagonalBorderOfCell.docx")

			' Get the first section of the document
			Dim section As Section = document.Sections(0)

			' Get the first table in the section
			Dim table As Table = TryCast(section.Tables(0), Table)

			' Create a StringBuilder to store the border information
			Dim stringBuilder As New StringBuilder()

			' Get the DiagonalUp border type of cell (0,0) in the table
			Dim bs_UP As Spire.Doc.Documents.BorderStyle = table(0, 0).CellFormat.Borders.DiagonalUp.BorderType
			stringBuilder.AppendLine("DiagonalUp border type of table cell (0,0) is " & bs_UP)

			' Get the DiagonalUp border color of cell (0,0) in the table
			Dim color_UP As Color = table(0, 0).CellFormat.Borders.DiagonalUp.Color
			stringBuilder.AppendLine("DiagonalUp border color of table cell (0,0) is ")
      stringBuilder.Append(color_UP)

			' Get the line width of the DiagonalUp border of cell (0,0) in the table
			Dim width_UP As Single = table(0, 0).CellFormat.Borders.DiagonalUp.LineWidth
			stringBuilder.AppendLine("Line width of DiagonalUp border of table cell (0,0) is " & width_UP)

			' Get the DiagonalDown border type of cell (0,0) in the table
			Dim bs_Down As Spire.Doc.Documents.BorderStyle = table(0, 0).CellFormat.Borders.DiagonalDown.BorderType
			stringBuilder.AppendLine("DiagonalDown border type of table cell (0,0) is " & bs_Down)

			' Get the DiagonalDown border color of cell (0,0) in the table
			Dim color_Down As Color = table(0, 0).CellFormat.Borders.DiagonalDown.Color
      stringBuilder.AppendLine("DiagonalDown border color of table cell (0,0) is ")
      stringBuilder.Append(color_Down)
 
			' Get the line width of the DiagonalDown border of cell (0,0) in the table
			Dim width_Down As Single = table(0, 0).CellFormat.Borders.DiagonalDown.LineWidth
			stringBuilder.AppendLine("DiagonalDown border line width of table cell (0,0) is " & width_Down)

			' Specify the output file path
			Dim output As String = "GetDiagonalBorder_out.txt"

			' Write the border information to the output file
			File.WriteAllText(output, stringBuilder.ToString())

			' Dispose of the document object to free up resources
			document.Dispose()

			'Launching the Word file.
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
