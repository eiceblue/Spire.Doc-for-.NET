Imports Spire.Doc
Imports Spire.Doc.Fields

Namespace AddPictureToTableCell
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

			'Get the first table from the first section of the document
			Dim table1 As Table = CType(doc.Sections(0).Tables(0), Table)

			'Add a picture to the specified table cell
			Dim picture As DocPicture = table1.Rows(1).Cells(2).Paragraphs(0).AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Spire.Doc.png"))
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim picture As DocPicture = table1.Rows(1).Cells(2).Paragraphs(0).AppendPicture("..\..\..\..\..\..\Data\Spire.Doc.png")
			' =============================================================================

			'Set picture width
			picture.Width = 100

			'Set picture height
			picture.Height = 100

			'Save the document
			Dim output As String = "AddPictureToTableCell.docx"
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
