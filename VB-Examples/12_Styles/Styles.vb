Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Interface

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

			'judge if it is Paragraph Style and then set paragraph format
			If TypeOf titleStyle Is ParagraphStyle Then
				Dim ps As ParagraphStyle = TryCast(titleStyle, ParagraphStyle)
				ps.CharacterFormat.Font = New System.Drawing.Font("cambria", 28)
				ps.CharacterFormat.TextColor = Color.FromArgb(42, 123, 136)
				ps.ParagraphFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single
				ps.ParagraphFormat.Borders.Bottom.Color = Color.FromArgb(42, 123, 136)
				ps.ParagraphFormat.Borders.Bottom.LineWidth = 1.5F
				ps.ParagraphFormat.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left
			End If
			'Add default normal style and modify
			Dim normalStyle As Style = document.AddStyle(BuiltinStyle.Normal)
			If TypeOf normalStyle Is ParagraphStyle Then
				Dim ps As ParagraphStyle = TryCast(normalStyle, ParagraphStyle)
				ps.CharacterFormat.Font = New System.Drawing.Font("cambria", 11)
			End If
			'Add default heading1 style
			Dim heading1Style As Style = document.AddStyle(BuiltinStyle.Heading1)
			If TypeOf heading1Style Is ParagraphStyle Then
				Dim ps As ParagraphStyle = TryCast(heading1Style, ParagraphStyle)
				ps.CharacterFormat.Font = New System.Drawing.Font("cambria", 14)
				ps.CharacterFormat.Bold = True
				ps.CharacterFormat.TextColor = Color.FromArgb(42, 123, 136)
			End If
			'Add default heading2 style
			Dim heading2Style As Style = document.AddStyle(BuiltinStyle.Heading2)
			If TypeOf heading2Style Is ParagraphStyle Then
				Dim ps As ParagraphStyle = TryCast(heading2Style, ParagraphStyle)
				ps.CharacterFormat.Font = New System.Drawing.Font("cambria", 12)
				ps.CharacterFormat.Bold = True
			End If

			'Create a bulleted list
			Dim bulletList As ListStyle = document.Styles.Add(ListType.Bulleted, "bulletList")
			If Not bulletList Is Nothing AndAlso TypeOf bulletList Is ICharacterStyle Then
				Dim style As ICharacterStyle = TryCast(bulletList, ICharacterStyle)
				style.CharacterFormat.Font = New System.Drawing.Font("cambria", 12)
			End If

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
			paragraph.ListFormat.ApplyStyle(bulletList)
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Minor:Text")
			paragraph.ListFormat.ApplyStyle(bulletList)
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Related coursework:Text")
			paragraph.ListFormat.ApplyStyle(bulletList)

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
			paragraph.ListFormat.ApplyStyle(bulletList)

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("LEADERSHIP")
			paragraph.ApplyStyle(BuiltinStyle.Heading2)

			'Add a paragraph and apply the style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Are you president of your fraternity, head of the condo board, or a team lead for your favorite charity? You¡¯re a natural leader¡ªtell it like it is!")
			paragraph.ListFormat.ApplyStyle(bulletList)

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
			paragraph.ListFormat.ApplyStyle(bulletList)

			'Save the document to a DOCX file
			Dim filePath As String = "Sample.docx"
			document.SaveToFile(filePath, FileFormat.Docx)

			'Dispose of the document object And open the created document in MS Word
			document.Dispose()
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
