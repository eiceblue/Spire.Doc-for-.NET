Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace PreventPageBreakInTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load an existing Word document from a file
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_Docx_5.docx")

			' Get the first table in the first section of the document
			Dim table As Table = TryCast(document.Sections(0).Tables(0), Table)

			' Iterate through each row in the table
			For Each row As TableRow In table.Rows
				' Iterate through each cell in the row
				For Each cell As TableCell In row.Cells
					' Iterate through each paragraph in the cell
					For Each p As Paragraph In cell.Paragraphs
						' Set "Keep with next" property to true to prevent page breaks within paragraphs
						p.Format.KeepFollow = True
					Next p
				Next cell
			Next row

			' Specify the output file path
			Dim result As String = "Result-PreventPageBreaksInWordTable.docx"

			' Save the modified document to a file
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
