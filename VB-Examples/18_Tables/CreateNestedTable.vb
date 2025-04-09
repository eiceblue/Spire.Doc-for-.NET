Imports Spire.Doc

Namespace CreateNestedTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a new document
			Dim doc As New Document()

			'Add a new section
			Dim section As Section = doc.AddSection()

			'Add a table
			Dim table As Table = section.AddTable(True)
			table.ResetCells(2, 2)

			'Set column width
			table.Rows(0).Cells(0).SetCellWidth(70, CellWidthType.Point)
			table.Rows(1).Cells(0).SetCellWidth(70, CellWidthType.Point)

			'Determines how Microsoft Word resizes a table when the AutoFit feature is used
			table.AutoFit(AutoFitBehaviorType.AutoFitToWindow)

			'Insert content to cells
			table(0, 0).AddParagraph().AppendText("Spire.Doc for .NET")
			Dim text As String = "Spire.Doc for .NET is a professional Word" & ".NET library specifically designed for developers to create," & "read, write, convert and print Word document files from any .NET" & "platform with fast and high quality performance."
			table(0, 1).AddParagraph().AppendText(text)

			'Add a nested table to cell(first row, second column)
			Dim nestedTable As Table = table(0, 1).AddTable(True)

			'Set the number of rows and columns
			nestedTable.ResetCells(4, 3)

			'Determines how Microsoft Word resizes a table when the AutoFit feature is used
			nestedTable.AutoFit(AutoFitBehaviorType.AutoFitToContents)

			'Add content to nested cells
			nestedTable(0, 0).AddParagraph().AppendText("NO.")
			nestedTable(0, 1).AddParagraph().AppendText("Item")
			nestedTable(0, 2).AddParagraph().AppendText("Price")

			'Add content to nested cells
			nestedTable(1, 0).AddParagraph().AppendText("1")
			nestedTable(1, 1).AddParagraph().AppendText("Pro Edition")
			nestedTable(1, 2).AddParagraph().AppendText("$799")

			'Add content to nested cells
			nestedTable(2, 0).AddParagraph().AppendText("2")
			nestedTable(2, 1).AddParagraph().AppendText("Standard Edition")
			nestedTable(2, 2).AddParagraph().AppendText("$599")

			'Add content to nested cells
			nestedTable(3, 0).AddParagraph().AppendText("3")
			nestedTable(3, 1).AddParagraph().AppendText("Free Edition")
			nestedTable(3, 2).AddParagraph().AppendText("$0")

			'Save the document
			Dim output As String = "CreateNestedTable.docx"
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
