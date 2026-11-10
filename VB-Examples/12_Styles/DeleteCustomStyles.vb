Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace DeleteCustomStyles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click

			' Create a new Document instance.
			Dim document As New Document()

			' Load an existing Word document that contains custom styles from the specified file path.
			document.LoadFromFile("..\..\..\..\..\..\..\Data\CustomStyles.docx")

			' Iterate through each section in the document.
			For Each section As Spire.Doc.Section In document.Sections
				' Iterate through each paragraph in the body of the current section.
				For Each paragraph As Paragraph In section.Body.Paragraphs
					' Remove the style associated with the current paragraph from the document's style collection.
					document.Styles(paragraph.StyleName).RemoveSelf()
				Next paragraph

				' Iterate through each table in the body of the current section.
				For Each table As Table In section.Body.Tables
					' Remove the style applied to the table itself.
					document.Styles(table.Format.StyleName).RemoveSelf()

					' Iterate through each row in the current table.
					For Each tableRow As Spire.Doc.TableRow In table.Rows
						' Iterate through each cell in the current table row.
						For Each tableCell As Spire.Doc.TableCell In tableRow.Cells
							' Iterate through each paragraph inside the current table cell.
							For Each cellParagraph As Paragraph In tableCell.Paragraphs
								' Remove the style associated with the paragraph inside the table cell.
								document.Styles(cellParagraph.StyleName).RemoveSelf()
							Next cellParagraph
						Next tableCell
					Next tableRow
				Next table
			Next section

			' Save the modified document to a new file with custom styles removed.
			document.SaveToFile("DeleteCustomStyles.docx", FileFormat.Docx)

			' Close the document to release internal resources.
			document.Close()

			' Explicitly dispose of the document object to free memory.
			document.Dispose()
			'Launching the Word file.
			WordDocViewer("DeleteCustomStyles.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
