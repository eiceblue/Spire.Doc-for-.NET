Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace AllowBreakAcrossPages
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a Word document
			Dim document As New Document()

			'Load the file from disk
			document.LoadFromFile("..\..\..\..\..\..\Data\AllowBreakAcrossPages.docx")

			'Get the first section
			Dim section As Section = document.Sections(0)

			'Get the first table
			Dim table As Table = TryCast(section.Tables(0), Table)

			'Loop through the table rows
			For Each row As TableRow In table.Rows

				'Allow break across pages
				row.RowFormat.IsBreakAcrossPages = True
			Next row

			'Save the Word document
			Dim output As String = "AllowBreakAcrossPages_out.docx"
			document.SaveToFile(output, FileFormat.Docx2013)

			'Dispose the document
			document.Dispose()

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
