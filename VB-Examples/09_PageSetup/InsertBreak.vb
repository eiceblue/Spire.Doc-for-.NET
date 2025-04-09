Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace InsertBreak
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Add a new section to the document
			Dim section As Section = document.AddSection()

			' Call the SetPage subroutine to set page size and margins for the section
			SetPage(section)

			' Call the InsertCover subroutine to insert cover content into the section
			InsertCover(section)

			' Add a new section and insert a section break to start a new page
			section = document.AddSection()
			section.AddParagraph().InsertSectionBreak(SectionBreakType.NewPage)

			' Call the InsertContent subroutine to insert content into the section
			InsertContent(section)

			' Save the document to a file with the specified file format
			document.SaveToFile("Sample.docx", FileFormat.Docx)

			' Dispose of the document object
			document.Dispose()

			'Launching the MS Word file.
			WordDocViewer("Sample.docx")
		End Sub

		' Subroutine to set page size and margins for the section
		Private Sub SetPage(ByVal section As Section)
			section.PageSetup.PageSize = PageSize.A4
			section.PageSetup.Margins.Top = 72.0F
			section.PageSetup.Margins.Bottom = 72.0F
			section.PageSetup.Margins.Left = 89.85F
			section.PageSetup.Margins.Right = 89.85F
		End Sub

		' Subroutine to insert cover content into the section
		Private Sub InsertCover(ByVal section As Section)
			' Create a new paragraph style for small text
			Dim small As New ParagraphStyle(section.Document)
			small.Name = "small"
			small.CharacterFormat.FontName = "Arial"
			small.CharacterFormat.FontSize = 9
			small.CharacterFormat.TextColor = Color.Gray
			section.Document.Styles.Add(small)

			' Add a paragraph with small text and apply the small style
			Dim paragraph As Paragraph = section.AddParagraph()
			paragraph.AppendText("The sample demonstrates how to insert section break.")
			paragraph.ApplyStyle(small.Name)

			' Add a title paragraph with large bold text and right alignment
			Dim title As Paragraph = section.AddParagraph()
			Dim text As TextRange = title.AppendText("Field Types Supported by Spire.Doc")
			text.CharacterFormat.FontName = "Arial"
			text.CharacterFormat.FontSize = 20
			text.CharacterFormat.Bold = True
			title.Format.BeforeSpacing = CInt(section.PageSetup.PageSize.Height) \ 2 - 3 * section.PageSetup.Margins.Top
			title.Format.AfterSpacing = 8
			title.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right

			' Add a paragraph with small text aligned to the right
			paragraph = section.AddParagraph()
			paragraph.AppendText("e-iceblue Spire.Doc team.")
			paragraph.ApplyStyle(small.Name)
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right
		End Sub

		' Subroutine to insert content into the section
		Private Sub InsertContent(ByVal section As Section)
			' Create a new paragraph style for the list
			Dim list As New ParagraphStyle(section.Document)
			list.Name = "list"
			list.CharacterFormat.FontName = "Arial"
			list.CharacterFormat.FontSize = 11
			list.ParagraphFormat.LineSpacing = 1.5F * 12.0F
			list.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple
			section.Document.Styles.Add(list)

			' Add title paragraph with the list style
			Dim title As Paragraph = section.AddParagraph()
			Dim text As TextRange = title.AppendText("Field type list:")
			title.ApplyStyle(list.Name)
			
			Dim first As Boolean = True

			' Iterate through each field type and add paragraphs with list numbering
			For Each type As FieldType In System.Enum.GetValues(GetType(FieldType))
				If type = FieldType.FieldUnknown OrElse type = FieldType.FieldNone OrElse type = FieldType.FieldEmpty Then
					Continue For
				End If
				Dim paragraph As Paragraph = section.AddParagraph()
				paragraph.AppendText(String.Format("{0} is supported in Spire.Doc", type))

				If first Then
					paragraph.ListFormat.ApplyNumberedStyle()
					first = False
				Else
					paragraph.ListFormat.ContinueListNumbering()
				End If
				paragraph.ApplyStyle(list.Name)
			Next type
		End Sub


		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
