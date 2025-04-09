Imports Spire.Doc

Namespace CloneRow
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

			'Get the first row of the first table
			Dim firstRow As TableRow = se.Tables(0).Rows(0)

			'Copy the first row to clone_FirstRow via TableRow.clone()
			Dim clone_FirstRow As TableRow = firstRow.Clone()

			'Add the table row to collection
			se.Tables(0).Rows.Add(clone_FirstRow)

			'Save the document
			Dim output As String = "CloneRow_output.docx"
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
