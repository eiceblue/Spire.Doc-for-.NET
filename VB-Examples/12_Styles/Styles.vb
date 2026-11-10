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
				' Initialize a new document instance
            Dim document As New Document()

            ' Add a new section to the document
            Dim sec As Section = document.AddSection()

            ' Add the built-in Title style to the document for modification
            Dim titleStyle As Style = document.AddStyle(BuiltinStyle.Title)

            ' Check if the added style is a ParagraphStyle, then configure its formatting
            If TypeOf titleStyle Is ParagraphStyle Then
                ' Cast the style to ParagraphStyle to access paragraph-specific properties
                Dim ps As ParagraphStyle = CType(titleStyle, ParagraphStyle)

                ' Set the font name, size, and text color for the title
                ps.CharacterFormat.FontName = "cambria"
                ps.CharacterFormat.FontSize = 28
                ps.CharacterFormat.TextColor = Color.FromArgb(42, 123, 136)

                ' Configure the bottom border of the title paragraph
                ps.ParagraphFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single
                ps.ParagraphFormat.Borders.Bottom.Color = Color.FromArgb(42, 123, 136)
                ps.ParagraphFormat.Borders.Bottom.LineWidth = 1.5F

                ' Align the title text to the left
                ps.ParagraphFormat.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left
            End If

            ' Add the built-in Normal style to the document for modification
            Dim normalStyle As Style = document.AddStyle(BuiltinStyle.Normal)

            ' Check if the added style is a ParagraphStyle, then configure its formatting
            If TypeOf normalStyle Is ParagraphStyle Then
                ' Cast the style to ParagraphStyle
                Dim ps As ParagraphStyle = CType(normalStyle, ParagraphStyle)

                ' Set the font name and size for normal text
                ps.CharacterFormat.FontName = "cambria"
                ps.CharacterFormat.FontSize = 11
            End If

            ' Add the built-in Heading1 style to the document
            Dim heading1Style As Style = document.AddStyle(BuiltinStyle.Heading1)

            ' Check if the added style is a ParagraphStyle, then configure its formatting
            If TypeOf heading1Style Is ParagraphStyle Then
                ' Cast the style to ParagraphStyle
                Dim ps As ParagraphStyle = CType(heading1Style, ParagraphStyle)

                ' Set the font name, size, bold weight, and text color for Heading1
                ps.CharacterFormat.FontName = "cambria"
                ps.CharacterFormat.FontSize = 14
                ps.CharacterFormat.Bold = True
                ps.CharacterFormat.TextColor = Color.FromArgb(42, 123, 136)
            End If

            ' Add the built-in Heading2 style to the document
            Dim heading2Style As Style = document.AddStyle(BuiltinStyle.Heading2)

            ' Check if the added style is a ParagraphStyle, then configure its formatting
            If TypeOf heading2Style Is ParagraphStyle Then
                ' Cast the style to ParagraphStyle
                Dim ps As ParagraphStyle = CType(heading2Style, ParagraphStyle)

                ' Set the font name, size, and bold weight for Heading2
                ps.CharacterFormat.FontName = "cambria"
                ps.CharacterFormat.FontSize = 12
                ps.CharacterFormat.Bold = True
            End If

            ' Add a custom bulleted list style named "bulletList"
            Dim bulletList As ListStyle = document.Styles.Add(ListType.Bulleted, "bulletList")

            ' Configure the bullet character and its formatting for the first list level
            bulletList.ListRef.Levels(0).BulletCharacter = ChrW(&H9F) ' Set bullet symbol to '?'
            bulletList.ListRef.Levels(0).CharacterFormat.FontName = "Wingdings" ' Use Wingdings font for the bullet
            bulletList.ListRef.Levels(0).CharacterFormat.FontSize = 12 ' Set bullet font size

            ' --- Apply styles to paragraphs in the section ---

            ' Add a new paragraph and apply the Title style
            Dim paragraph As Spire.Doc.Documents.Paragraph = sec.AddParagraph()
            paragraph.AppendText("Your Name")
            paragraph.ApplyStyle(BuiltinStyle.Title)

            ' Add a new paragraph for contact info and apply the Normal style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("Address, City, ST ZIP Code | Telephone | Email")
            paragraph.ApplyStyle(BuiltinStyle.Normal)

            ' Add a new paragraph for the "Objective" heading and apply Heading1 style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("Objective")
            paragraph.ApplyStyle(BuiltinStyle.Heading1)

            ' Add a new paragraph for the objective description and apply Normal style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("To get started right away, just click any placeholder text (such as this) and start typing to replace it with your own.")
            paragraph.ApplyStyle(BuiltinStyle.Normal)

            ' Add a new paragraph for the "Education" heading and apply Heading1 style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("Education")
            paragraph.ApplyStyle(BuiltinStyle.Heading1)

            ' Add a new paragraph for the education subheading and apply Heading2 style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("DEGREE | DATE EARNED | SCHOOL")
            paragraph.ApplyStyle(BuiltinStyle.Heading2)

            ' Add a paragraph for the Major and apply the custom bullet list style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("Major:Text")
            paragraph.ListFormat.ApplyStyle(bulletList)

            ' Add a paragraph for the Minor and apply the custom bullet list style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("Minor:Text")
            paragraph.ListFormat.ApplyStyle(bulletList)

            ' Add a paragraph for Related Coursework and apply the custom bullet list style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("Related coursework:Text")
            paragraph.ListFormat.ApplyStyle(bulletList)

            ' Add a new section heading "Skills & Abilities" and apply the Heading1 style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("Skills & Abilities")
            paragraph.ApplyStyle(BuiltinStyle.Heading1)

            ' Add a subheading for "MANAGEMENT" skills and apply the Heading2 style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("MANAGEMENT")
            paragraph.ApplyStyle(BuiltinStyle.Heading2)

            ' Add a bulleted description for the Management section
            paragraph = sec.AddParagraph()
            paragraph.AppendText("Think a document that looks this good has to be difficult to format? Think again! To easily apply any text formatting you see in this document with just a click, on the Home tab of the ribbon, check out Styles.")
            paragraph.ListFormat.ApplyStyle(bulletList)

            ' Add a subheading for "COMMUNICATION" skills and apply the Heading2 style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("COMMUNICATION")
            paragraph.ApplyStyle(BuiltinStyle.Heading2)

            ' Add a bulleted description for the Communication section
            paragraph = sec.AddParagraph()
            paragraph.AppendText("You delivered that big presentation to rave reviews. Don't be shy about it now! This is the place to show how well you work and play with others.")
            paragraph.ListFormat.ApplyStyle(bulletList)

            ' Add a subheading for "LEADERSHIP" skills and apply the Heading2 style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("LEADERSHIP")
            paragraph.ApplyStyle(BuiltinStyle.Heading2)

            ' Add a bulleted description for the Leadership section
            paragraph = sec.AddParagraph()
            paragraph.AppendText("Are you president of your fraternity, head of the condo board, or a team lead for your favorite charity? You're a natural leader¡ªtell it like it is!")
            paragraph.ListFormat.ApplyStyle(bulletList)

            ' Add a new section heading "Experience" and apply the Heading1 style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("Experience")
            paragraph.ApplyStyle(BuiltinStyle.Heading1)

            ' Add a subheading for the job title placeholder and apply the Heading2 style
            paragraph = sec.AddParagraph()
            paragraph.AppendText("JOB TITLE | COMPANY | DATES FROM - TO")
            paragraph.ApplyStyle(BuiltinStyle.Heading2)

            ' Add a bulleted description for the Experience section
            paragraph = sec.AddParagraph()
            paragraph.AppendText("This is the place for a brief summary of your key responsibilities and most stellar accomplishments.")
            paragraph.ListFormat.ApplyStyle(bulletList)

            ' Save the generated document to a .docx file
            Dim filePath As String = "Sample.docx"
            document.SaveToFile(filePath, Spire.Doc.FileFormat.Docx)

            ' Close the document to release resources
            document.Close()

            ' Dispose of the document object to free up memory
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
