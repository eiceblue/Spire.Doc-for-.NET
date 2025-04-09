Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace CloneTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			Dim input As String = "..\..\..\..\..\..\Data\TableTemplate.docx"
			
			'Create a Word document
			Dim doc As New Document()
			
			'Load the file from disk
			doc.LoadFromFile(input)

			'Get the first section
			Dim se As Section = doc.Sections(0)

			'Get the first table
			Dim original_Table As Table = CType(se.Tables(0), Table)

			'Copy the existing table to copied_Table via Table.clone()
			Dim copied_Table As Table = original_Table.Clone()
			Dim st() As String = { "Spire.Presentation for .Net", "A professional " & "PowerPoint® compatible library that enables developers to create, read, " & "write, modify, convert and Print PowerPoint documents on any .NET framework, " & ".NET Core platform." }
			'Get the last row of table
			Dim lastRow As TableRow = copied_Table.Rows(copied_Table.Rows.Count - 1)
			'Change last row data
			For i As Integer = 0 To lastRow.Cells.Count - 2
				lastRow.Cells(i).Paragraphs(0).Text = st(i)
			Next i
			'Add copied_Table in section
			se.Tables.Add(copied_Table)

			'Save and launch document
			Dim output As String = "CloneTable.docx"
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
