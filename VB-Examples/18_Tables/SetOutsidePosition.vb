Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace SetOutsidePosition
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
		   ' Create a new document object
		   Dim doc As New Document()

		   ' Add a section to the document
		   Dim sec As Section = doc.AddSection()

		   ' Get the header of the first section in the document
		   Dim header As HeaderFooter = doc.Sections(0).HeadersFooters.Header

		   ' Add a paragraph to the header with left-aligned text
		   Dim paragraph As Paragraph = header.AddParagraph()
		   paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left

		   ' Append an image to the paragraph in the header
		   Dim headerimage As DocPicture = paragraph.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Word.png"))

		   ' Add a table to the header
		   Dim table As Table = header.AddTable()
		   table.ResetCells(4, 2)

		   ' Set table properties for text wrapping and positioning
		   table.Format.WrapTextAround = True
		   table.Format.Positioning.HorizPositionAbs = HorizontalPosition.Outside
		   table.Format.Positioning.VertRelationTo = VerticalRelation.Margin
		   table.Format.Positioning.VertPosition = 43

		   ' Define data for the table cells
		   Dim data()() As String = { New String() {"Spire.Doc.left", "Spire XLS.right"}, New String() {"Spire.Presentatio.left", "Spire.PDF.right"}, New String() {"Spire.DataExport.left", "Spire.PDFViewe.right"}, New String() {"Spire.DocViewer.left", "Spire.BarCode.right"} }

		   ' Fill the table with data and set cell widths
		   For r As Integer = 0 To 3
			   Dim dataRow As TableRow = table.Rows(r)
			   For c As Integer = 0 To 1
				   If c = 0 Then
					   ' Add left-aligned text to the cell
					   Dim par As Paragraph = dataRow.Cells(c).AddParagraph()
					   par.AppendText(data(r)(c))
					   par.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left
					   dataRow.Cells(c).SetCellWidth(180, CellWidthType.Point)
				   Else
					   ' Add right-aligned text to the cell
					   Dim par As Paragraph = dataRow.Cells(c).AddParagraph()
					   par.AppendText(data(r)(c))
					   par.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right
					   dataRow.Cells(c).SetCellWidth(180, CellWidthType.Point)
				   End If
			   Next c
		   Next r

		   ' Specify the output file name
		   Dim output As String = "SetOutsidePosition.docx"

		   ' Save the document to a file in Docx format
		   doc.SaveToFile(output, FileFormat.Docx)

		   ' Dispose of the document object
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
