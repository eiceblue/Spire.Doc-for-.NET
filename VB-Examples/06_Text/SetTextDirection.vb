Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace SetTextDirection
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object.
			Dim doc As New Document()

			' Add a section to the document.
			Dim section1 As Section = doc.AddSection()

			' Set the text direction of section1 to right-to-left.
			section1.TextDirection = TextDirection.RightToLeft

			' Create a new ParagraphStyle and set its properties.
			Dim style As New ParagraphStyle(doc)
			style.Name = "FontStyle"
			style.CharacterFormat.FontName = "Arial"
			style.CharacterFormat.FontSize = 15

			' Add the style to the document's styles collection.
			doc.Styles.Add(style)

			' Add a paragraph to section1 and set its text content.
			Dim p As Paragraph = section1.AddParagraph()
			p.AppendText("Only Spire.Doc, no Microsoft Office automation")

			' Apply the previously created style to the paragraph.
			p.ApplyStyle(style.Name)

			' Add another paragraph to section1 with different text content.
			p = section1.AddParagraph()
			p.AppendText("Convert file documents with high quality")

			' Apply the same style to the second paragraph.
			p.ApplyStyle(style.Name)

			' Add another section to the document.
			Dim section2 As Section = doc.AddSection()

			' Add a table to section2.
			Dim table As Table = section2.AddTable()
			table.ResetCells(1, 1)

			' Get a reference to the only cell in the table.
			Dim cell As TableCell = table.Rows(0).Cells(0)

			' Set the height of the table row.
			table.Rows(0).Height = 150

			' Set the width of the cell.
			table.Rows(0).Cells(0).SetCellWidth(10, CellWidthType.Point)

			' Set the text direction of the cell to right-to-left rotated.
			cell.CellFormat.TextDirection = TextDirection.RightToLeftRotated

			' Add a paragraph to the cell and set its text content.
			cell.AddParagraph().AppendText("This is vertical style")

			' Add another paragraph to section2 with different text content.
			p = section2.AddParagraph()
			p.AppendText("This is horizontal style")

			' Apply the same style to the second paragraph.
			p.ApplyStyle(style.Name)

			' Save the document to a file named "SetTextDirection.docx".
			Dim output As String = "SetTextDirection.docx"
			doc.SaveToFile(output, FileFormat.Docx)

			' Dispose of the document object to free up resources.
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
