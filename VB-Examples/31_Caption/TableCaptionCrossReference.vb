Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields
Imports Spire.Doc.Interface

Namespace TableCaptionCrossReference
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of Document
			Dim document As New Document()

			' Add a section to the document
			Dim section As Section = document.AddSection()

			' Add a table to the section with 2 rows and 3 columns
			Dim table As Table = section.AddTable(True)
			table.ResetCells(2, 3)

			' Add a caption to the table with the "Table" label, numbering format as "Number", and position below the table
			Dim captionParagraph As IParagraph = table.AddCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.BelowItem)

			' Add a bookmark at the specified location
			Dim bookmarkName As String = "Table_1"
			Dim paragraph As Paragraph = section.AddParagraph()
			paragraph.AppendBookmarkStart(bookmarkName)
			paragraph.AppendBookmarkEnd(bookmarkName)

			' Navigate to the bookmark and replace its content with the caption paragraph
			Dim navigator As New BookmarksNavigator(document)
			navigator.MoveToBookmark(bookmarkName)
			Dim part As TextBodyPart = navigator.GetBookmarkContent()
			part.BodyItems.Clear()
			part.BodyItems.Add(captionParagraph)
			navigator.ReplaceBookmarkContent(part)

			' Create a cross-reference field for the bookmark
			Dim field As New Field(document)
			field.Type = FieldType.FieldRef
			field.Code = "REF Table_1 \p \h"

			' Add line breaks before the next paragraph
			For i As Integer = 0 To 2
				paragraph.AppendBreak(BreakType.LineBreak)
			Next i

			' Add a new paragraph for the caption cross-reference
			paragraph = section.AddParagraph()

			' Add text to the paragraph
			Dim range As TextRange = paragraph.AppendText("This is a table caption cross-reference, ")
			range.CharacterFormat.FontSize = 14

			' Add the field for referencing the table caption
			paragraph.ChildObjects.Add(field)

			' Add a field separator
			Dim fieldSeparator As New FieldMark(document, FieldMarkType.FieldSeparator)
			paragraph.ChildObjects.Add(fieldSeparator)

			' Add the text "Table 1" as the reference text
			Dim tr As New TextRange(document)
			tr.Text = "Table 1"
			tr.CharacterFormat.FontSize = 14
			tr.CharacterFormat.TextColor = Color.DeepSkyBlue
			paragraph.ChildObjects.Add(tr)

			' Add a field end mark
			Dim fieldEnd As New FieldMark(document, FieldMarkType.FieldEnd)
			paragraph.ChildObjects.Add(fieldEnd)

			' Enable field updating in the document
			document.IsUpdateFields = True

			' Specify the output file name and format (Docx)
			Dim output As String = "TableCaptionCrossReference.docx"
			document.SaveToFile(output, FileFormat.Docx)

			' Dispose of the document object when finished using it
			document.Dispose()

			'Launching the file
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
