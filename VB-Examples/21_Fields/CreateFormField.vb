Imports System.IO
Imports System.Xml.XPath
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace CreateFormField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Add a section to the document
			Dim section As Section = document.AddSection()

			' Set page settings for the section
			SetPage(section)

			' Insert header and footer into the section
			InsertHeaderAndFooter(section)

			' Add a title to the section
			AddTitle(section)

			' Add a form to the section
			AddForm(section)

			' Save the document to a file with the specified format
			document.SaveToFile("Sample.doc", FileFormat.Doc)

			' Dispose the document object
			document.Dispose()

			'Launch the Word file.
			WordDocViewer("Sample.doc")


		End Sub

		Private Shared Sub SetPage(ByVal section As Section)
			' Set the page size of the section to A4
			section.PageSetup.PageSize = PageSize.A4

			' Set the top, bottom, left, and right margins of the section
			section.PageSetup.Margins.Top = 72f
			section.PageSetup.Margins.Bottom = 72f
			section.PageSetup.Margins.Left = 89.85f
			section.PageSetup.Margins.Right = 89.85f
		End Sub

		Private Shared Sub InsertHeaderAndFooter(ByVal section As Section)
			' Add a paragraph to the header and insert a picture
			Dim headerParagraph As Paragraph = section.HeadersFooters.Header.AddParagraph()
			Dim headerPicture As DocPicture = headerParagraph.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Header.png"))
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim headerPicture As DocPicture = headerParagraph.AppendPicture("..\..\..\..\..\..\Data\Header.png")
			' =============================================================================

			' Add text to the header paragraph with specified font settings
			Dim text As TextRange = headerParagraph.AppendText("Demo of Spire.Doc")
			text.CharacterFormat.FontName = "Arial"
			text.CharacterFormat.FontSize = 10
			text.CharacterFormat.Italic = True
			headerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right

			' Set border settings for the bottom border of the header paragraph
			headerParagraph.Format.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single
			headerParagraph.Format.Borders.Bottom.Space = 0.05F

			' Set wrapping style and alignment for the header picture
			headerPicture.TextWrappingStyle = TextWrappingStyle.Behind
			headerPicture.HorizontalOrigin = HorizontalOrigin.Page
			headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
			headerPicture.VerticalOrigin = VerticalOrigin.Page
			headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top

			' Add a paragraph to the footer and insert a picture
			Dim footerParagraph As Paragraph = section.HeadersFooters.Footer.AddParagraph()
			Dim footerPicture As DocPicture = footerParagraph.AppendPicture(Image.FromFile("..\..\..\..\..\..\Data\Footer.png"))
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'Dim footerPicture As DocPicture = footerParagraph.AppendPicture("..\..\..\..\..\..\Data\Footer.png")
			' =============================================================================

			' Set wrapping style and alignment for the footer picture
			footerPicture.TextWrappingStyle = TextWrappingStyle.Behind
			footerPicture.HorizontalOrigin = HorizontalOrigin.Page
			footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left
			footerPicture.VerticalOrigin = VerticalOrigin.Page
			footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom

			' Append field codes for page number and number of pages to the footer paragraph
			footerParagraph.AppendField("page number", FieldType.FieldPage)
			footerParagraph.AppendText(" of ")
			footerParagraph.AppendField("number of pages", FieldType.FieldNumPages)
			footerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right

			' Set border settings for the top border of the footer paragraph
			footerParagraph.Format.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single
			footerParagraph.Format.Borders.Top.Space = 0.05F
		End Sub

		Private Shared Sub AddTitle(ByVal section As Section)
			' Add a paragraph for the title
			Dim title As Paragraph = section.AddParagraph()

			' Append the title text with specified font settings
			Dim titleText As TextRange = title.AppendText("Create Your Account")
			titleText.CharacterFormat.FontSize = 18
			titleText.CharacterFormat.FontName = "Arial"
			titleText.CharacterFormat.TextColor = Color.FromArgb(&H0, &H71, &Hb6)

			' Set the horizontal alignment of the title paragraph to center
			title.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center

			' Set the spacing after the title paragraph
			title.Format.AfterSpacing = 8
		End Sub

		Private Shared Sub AddForm(ByVal section As Section)
			' Create a paragraph style for description texts
			Dim descriptionStyle As New ParagraphStyle(section.Document)
			descriptionStyle.Name = "description"
			descriptionStyle.CharacterFormat.FontSize = 12
			descriptionStyle.CharacterFormat.FontName = "Arial"
			descriptionStyle.CharacterFormat.TextColor = Color.FromArgb(&H0, &H45, &H8e)
			section.Document.Styles.Add(descriptionStyle)

			' Add the first description paragraph
			Dim p1 As Paragraph = section.AddParagraph()
			Dim text1 As String = "So that we can verify your identity and find your information, " & "please provide us with the following information. " & "This information will be used to create your online account. " & "Your information is not public, shared in any way, or displayed on this site"
			p1.AppendText(text1)
			p1.ApplyStyle(descriptionStyle.Name)

			' Add the second description paragraph
			Dim p2 As Paragraph = section.AddParagraph()
			Dim text2 As String = "You must provide a real email address to which we will send your password."
			p2.AppendText(text2)
			p2.ApplyStyle(descriptionStyle.Name)
			p2.Format.AfterSpacing = 8

			' Create a paragraph style for form field group labels
			Dim formFieldGroupLabelStyle As New ParagraphStyle(section.Document)
			formFieldGroupLabelStyle.Name = "formFieldGroupLabel"
			formFieldGroupLabelStyle.ApplyBaseStyle("description")
			formFieldGroupLabelStyle.CharacterFormat.Bold = True
			formFieldGroupLabelStyle.CharacterFormat.TextColor = Color.White
			section.Document.Styles.Add(formFieldGroupLabelStyle)

			' Create a paragraph style for form field labels
			Dim formFieldLabelStyle As New ParagraphStyle(section.Document)
			formFieldLabelStyle.Name = "formFieldLabel"
			formFieldLabelStyle.ApplyBaseStyle("description")
			formFieldLabelStyle.ParagraphFormat.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right
			section.Document.Styles.Add(formFieldLabelStyle)

			' Add a table to the section for the form fields
			Dim table As Table = section.AddTable()
			table.DefaultColumnsNumber = 2 ' Set the number of columns
			table.DefaultRowHeight = 20 ' Set the default row height

			' Read the XML file containing the form structure
			Using stream As Stream = File.OpenRead("..\..\..\..\..\..\Data\Form.xml")
				Dim xpathDoc As New XPathDocument(stream)
				Dim sectionNodes As XPathNodeIterator = xpathDoc.CreateNavigator().Select("/form/section")

				' Iterate over each section node in the XML file
				For Each node As XPathNavigator In sectionNodes
					' Add a row for the form field group label
					Dim row As TableRow = table.AddRow(False)
					row.Cells(0).CellFormat.Shading.BackgroundPatternColor= Color.FromArgb(&H0, &H71, &Hb6)
					row.Cells(0).CellFormat.VerticalAlignment = VerticalAlignment.Middle

					' Add the form field group label text to the cell
					Dim cellParagraph As Paragraph = row.Cells(0).AddParagraph()
					cellParagraph.AppendText(node.GetAttribute("name", ""))
					cellParagraph.ApplyStyle(formFieldGroupLabelStyle.Name)

					' Iterate over each field node within the section node
					Dim fieldNodes As XPathNodeIterator = node.Select("field")
					For Each fieldNode As XPathNavigator In fieldNodes
						' Add a row for the form field label and input field
						Dim fieldRow As TableRow = table.AddRow(False)

						' Set vertical alignment for the cells in the field row
						fieldRow.Cells(0).CellFormat.VerticalAlignment = VerticalAlignment.Middle
						fieldRow.Cells(1).CellFormat.VerticalAlignment = VerticalAlignment.Middle

						' Add the form field label to the first cell in the row
						Dim labelParagraph As Paragraph = fieldRow.Cells(0).AddParagraph()
						labelParagraph.AppendText(fieldNode.GetAttribute("label", ""))
						labelParagraph.ApplyStyle(formFieldLabelStyle.Name)

						' Add the input field paragraph to the second cell in the row
						Dim fieldParagraph As Paragraph = fieldRow.Cells(1).AddParagraph()
						Dim fieldId As String = fieldNode.GetAttribute("id", "")
						Select Case fieldNode.GetAttribute("type", "")
							Case "text"
								' Add a text form input field
								Dim field As TextFormField = TryCast(fieldParagraph.AppendField(fieldId, FieldType.FieldFormTextInput), TextFormField)
								field.DefaultText = ""
								field.Text = ""

							Case "list"
								' Add a dropdown list form field
								Dim list As DropDownFormField = TryCast(fieldParagraph.AppendField(fieldId, FieldType.FieldFormDropDown), DropDownFormField)


								Dim itemNodes As XPathNodeIterator = fieldNode.Select("item")
								For Each itemNode As XPathNavigator In itemNodes
									list.DropDownItems.Add(itemNode.SelectSingleNode("text()").Value)
								Next itemNode

							Case "checkbox"
								' Add a checkbox form field
								fieldParagraph.AppendField(fieldId, FieldType.FieldFormCheckBox)
						End Select
					Next fieldNode


					table.ApplyHorizontalMerge(row.GetRowIndex(), 0, 1)
				Next node
			End Using
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
