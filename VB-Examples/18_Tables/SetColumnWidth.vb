Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc

Namespace SetColumnWidth
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim document As New Document()

			' Load a document from a file, specified by the file path
			document.LoadFromFile("..\..\..\..\..\..\Data\TableSample.docx")

			' Get the first section of the document
			Dim section As Section = document.Sections(0)

			' Get the first table in the section
			Dim table As Table = TryCast(section.Tables(0), Table)

			' Set the width of the first column in each row to 200 points
			For i As Integer = 0 To table.Rows.Count - 1
				table.Rows(i).Cells(0).SetCellWidth(200, CellWidthType.Point)
			Next i

			' Save the modified document to a file with the name "ColumnWidth.docx", using Docx format
			document.SaveToFile("ColumnWidth.docx", FileFormat.Docx)

			' Dispose of the document object
			document.Dispose() 
			
			'Launch the document
			FileViewer("ColumnWidth.docx")
		End Sub
		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
