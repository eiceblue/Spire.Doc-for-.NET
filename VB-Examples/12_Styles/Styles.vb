Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace Styles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
				'Create a Word document
			Dim document As New Document()

			'Add a section
			Dim sec As Section = document.AddSection()

			'Add default title style to document and set its format
			Dim titleStyle As Style = document.AddStyle(BuiltinStyle.Title)
			titleStyle.CharacterFormat.Font = New Font("cambria", 28)
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'titleStyle.CharacterFormat.FontName= "cambria";
			'titleStyle.CharacterFormat.FontSize = 28;
			' =============================================================================
			titleStyle.CharacterFormat.TextColor = Color.FromArgb(42, 123, 136)
			'judge if it is Paragraph Style and then set paragraph format
			If TypeOf titleStyle Is ParagraphStyle Then
				Dim ps As ParagraphStyle = TryCast(titleStyle, ParagraphStyle)
				ps.ParagraphFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single
				ps.ParagraphFormat.Borders.Bottom.Color = Color.FromArgb(42, 123, 136)
				ps.ParagraphFormat.Borders.Bottom.LineWidth = 1.5F
				ps.ParagraphFormat.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left
			End If
			'Add default normal style and modify
			Dim normalStyle As Style = document.AddStyle(BuiltinStyle.Normal)
			normalStyle.CharacterFormat.Font = New Font("cambria", 11)
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'normalStyle.CharacterFormat.FontName = "cambria";
			'normalStyle.CharacterFormat.FontSize = 11;
			' =============================================================================
			'Add default heading1 style
			Dim heading1Style As Style = document.AddStyle(BuiltinStyle.Heading1)
			heading1Style.CharacterFormat.Font = New Font("cambria", 14)
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'heading1Style.CharacterFormat.FontName = "cambria";
			'heading1Style.CharacterFormat.FontSize = 14;
			' =============================================================================
			heading1Style.CharacterFormat.Bold = True
			heading1Style.CharacterFormat.TextColor = Color.FromArgb(42, 123, 136)
			'Add default heading2 style
			Dim heading2Style As Style = document.AddStyle(BuiltinStyle.Heading2)
			heading2Style.CharacterFormat.Font = New Font("cambria", 12)
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'heading2Style.CharacterFormat.FontName = "cambria";
			'heading2Style.CharacterFormat.FontSize = 12;
			' =============================================================================
			heading2Style.CharacterFormat.Bold = True

			'Create a bulleted list
			Dim bulletList As New ListStyle(document, ListType.Bulleted)
			bulletList.CharacterFormat.Font = New Font("cambria", 12)
			' =============================================================================
			' Use the following code for netstandard dlls
			' =============================================================================
			'bulletList.CharacterFormat.FontName = "cambria";
			'bulletList.CharacterFormat.FontSize = 12;
			' =============================================================================
			bulletList.Name = "bulletList"

			'Add the list to the document
			document.ListStyles.Add(bulletList)


			'Add a paragraph and apply the style
			Dim paragraph As Paragraph = sec.AddParagraph()
			paragraph.AppendText("Your Name")
			paragraph.ApplyStyle(BuiltinStyle.Title)

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Address, City, ST ZIP Code | Telephone | Email")
			paragraph.ApplyStyle(BuiltinStyle.Normal)

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Objective")
			paragraph.ApplyStyle(BuiltinStyle.Heading1)

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("To get started right away, just click any placeholder text (such as this) and start typing to replace it with your own.")
			paragraph.ApplyStyle(BuiltinStyle.Normal)

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Education")
			paragraph.ApplyStyle(BuiltinStyle.Heading1)

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("DEGREE | DATE EARNED | SCHOOL")
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			'Add a paragraph and apply the style named bulletList
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Major:Text")
			paragraph.ListFormat.ApplyStyle("bulletList")
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Minor:Text")
			paragraph.ListFormat.ApplyStyle("bulletList")
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Related coursework:Text")
			paragraph.ListFormat.ApplyStyle("bulletList")

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Skills & Abilities")
			paragraph.ApplyStyle(BuiltinStyle.Heading1)

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("MANAGEMENT")
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Think a document that looks this good has to be difficult to format? Think again! To easily apply any text formatting you see in this document with just a click, on the Home tab of the ribbon, check out Styles.")
			paragraph.ListFormat.ApplyStyle("bulletList")

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("COMMUNICATION")
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("You delivered that big presentation to rave reviews. Don¡¯t be shy about it now! This is the place to show how well you work and play with others.")
			paragraph.ListFormat.ApplyStyle("bulletList")

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("LEADERSHIP")
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Are you president of your fraternity, head of the condo board, or a team lead for your favorite charity? You¡¯re a natural leader¡ªtell it like it is!")
			paragraph.ListFormat.ApplyStyle("bulletList")

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Experience")
			paragraph.ApplyStyle(BuiltinStyle.Heading1)

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("JOB TITLE | COMPANY | DATES FROM - TO")
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("This is the place for a brief summary of your key responsibilities and most stellar accomplishments.")
			paragraph.ListFormat.ApplyStyle("bulletList")

			'Save to docx file.
			Dim filePath As String = "Sample.docx"
			document.SaveToFile(filePath, FileFormat.Docx)

			'Dispose of the document object
			document.Dispose()
			'Launching the MS Word file.
			WordDocViewer(filePath)
		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
