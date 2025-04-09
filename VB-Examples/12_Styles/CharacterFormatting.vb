Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace FontAndColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document object
			Dim document As New Document()

			' Add a new section to the document
			Dim sec As Section = document.AddSection()

			' Add a title paragraph to the section
			Dim titleParagraph As Paragraph = sec.AddParagraph()
			titleParagraph.AppendText("Font Styles and Effects ")
			titleParagraph.ApplyStyle(BuiltinStyle.Title)

			' Add a regular paragraph to the section
			Dim paragraph As Paragraph = sec.AddParagraph()

			' Add strikethrough text to the paragraph
			Dim tr As TextRange = paragraph.AppendText("Strikethough Text")
			tr.CharacterFormat.IsStrikeout = True

			' Add a line break and shadow text to the paragraph
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Shadow Text")
			tr.CharacterFormat.IsShadow = True

			' Add a line break and small caps text to the paragraph
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Small caps Text")
			tr.CharacterFormat.IsSmallCaps = True

			' Add a line break and double strikethough text to the paragraph
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Double Strikethough Text")
			tr.CharacterFormat.DoubleStrike = True

			' Add a line break and outline text to the paragraph
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Outline Text")
			tr.CharacterFormat.IsOutLine = True

			' Add a line break and all caps text to the paragraph
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("AllCaps Text")
			tr.CharacterFormat.AllCaps = True

			' Add a line break and subscript and superscript text to the paragraph
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Text")
			tr = paragraph.AppendText("SubScript")
			tr.CharacterFormat.SubSuperScript = SubSuperScript.SubScript

			tr = paragraph.AppendText("And")
			tr = paragraph.AppendText("SuperScript")
			tr.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript

			' Add a line break and emboss text to the paragraph
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Emboss Text")
			tr.CharacterFormat.Emboss = True
			tr.CharacterFormat.TextColor = Color.White

			' Add a line break and hidden text to the paragraph
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Hidden:")
			tr = paragraph.AppendText("Hidden Text")
			tr.CharacterFormat.Hidden = True

			' Add a line break and engrave text to the paragraph
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Engrave Text")
			tr.CharacterFormat.Engrave = True
			tr.CharacterFormat.TextColor = Color.White

			' Add a line break and set different font names for different character sets
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("WesternFonts╓╨╬─╫╓╠х")
			tr.CharacterFormat.FontNameAscii = "Calibri"
			tr.CharacterFormat.FontNameNonFarEast = "Calibri"
			tr.CharacterFormat.FontNameFarEast = "Simsun"

			' Add a line break and set the font size for the text
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Font Size")
			tr.CharacterFormat.FontSize = 20

			' Add a line break and set the font color for the text
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Font Color")
			tr.CharacterFormat.TextColor = Color.Red

			' Add a line break and set the bold and italic styles for the text
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Bold Italic Text")
			tr.CharacterFormat.Bold = True
			tr.CharacterFormat.Italic = True

			' Add a line break and set the underline style for the text
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Underline Style")
			tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single

			' Add a line break and set the highlight color for the text
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Highlight Text")
			tr.CharacterFormat.HighlightColor = Color.Yellow

			' Add a line break and set the text background color for the text
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Text has shading")
			tr.CharacterFormat.TextBackgroundColor = Color.Green

			' Add a line break and set a border around the text
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Border Around Text")
			tr.CharacterFormat.Border.BorderType = Spire.Doc.Documents.BorderStyle.Single

			' Add a line break and set the text scale for the text
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Text Scale")
			tr.CharacterFormat.TextScale = 150

			' Add a line break and set the character spacing for the text
			paragraph.AppendBreak(BreakType.LineBreak)
			tr = paragraph.AppendText("Character Spacing is 2 point")
			tr.CharacterFormat.CharacterSpacing = 2

			' Save the document to a file
			Dim filePath As String = "CharaterFormatting.docx"
			document.SaveToFile(filePath, FileFormat.Docx)
			
			' Dispose of the document object
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
