Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace CreateVerticalTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Create a new Document object
			Dim document As New Document()

			' Add a new section to the document
			Dim section As Section = document.AddSection()

			' Add a table to the section
			Dim table As Table = section.AddTable()
			table.ResetCells(1, 1)

			' Get the first cell of the table
			Dim cell As TableCell = table.Rows(0).Cells(0)

			' Set the height of the table row
			table.Rows(0).Height = 150

			' Add a paragraph with text to the cell
			cell.AddParagraph().AppendText("Draft copy in vertical style")

			' Set the text direction of the cell to right-to-left rotated
			cell.CellFormat.TextDirection = TextDirection.RightToLeftRotated

			' Enable wrap text around the table
			table.Format.WrapTextAround = True

			' Set the vertical position of the table relative to the page
			table.Format.Positioning.VertRelationTo = VerticalRelation.Page

			' Set the horizontal position of the table relative to the page
			table.Format.Positioning.HorizRelationTo = HorizontalRelation.Page

			' Set the horizontal position of the table
			table.Format.Positioning.HorizPosition = section.PageSetup.PageSize.Width - table.Width

			' Set the vertical position of the table
			table.Format.Positioning.VertPosition = 200

			' Save the document to a file in Docx2013 format
			Dim result As String = "Result-CreateVerticalTable.docx"
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document object to free up resources
			document.Dispose()

			'Launch the MS Word file.
			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
