Imports Spire.Doc
Imports Spire.Doc.Collections
Imports Spire.Doc.Documents

Namespace CompareListLevels
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of a Word document.
			Dim document As New Document()

			' Add a new bulleted list style named "bulletList1" to the document's styles collection.
			Dim listStyle_1 As ListStyle = document.Styles.Add(ListType.Bulleted, "bulletList1")

			' Retrieve the list level collection associated with the first list style.
			Dim Levels_1 As ListLevelCollection = listStyle_1.ListRef.Levels

			' Get the first (top-level) list formatting object from the first style.
			Dim L10 As ListLevel = Levels_1(0)

			' Get the second-level list formatting object from the first style.
			Dim L11 As ListLevel = Levels_1(1)

			' Get the third-level list formatting object from the first style.
			Dim L12 As ListLevel = Levels_1(2)

			' Add another bulleted list style named "bulletList2" to the document.
			Dim listStyle_2 As ListStyle = document.Styles.Add(ListType.Bulleted, "bulletList2")

			' Retrieve the list level collection for the second list style.
			Dim Levels_2 As ListLevelCollection = listStyle_2.ListRef.Levels

			' Get the first-level list formatting object from the second style.
			Dim L20 As ListLevel = Levels_2(0)

			' Get the second-level list formatting object from the second style.
			Dim L21 As ListLevel = Levels_2(1)

			' Get the third-level list formatting object from the second style.
			Dim L22 As ListLevel = Levels_2(2)

			' Set line spacing for the first level of the first list style (1.5 times 10-point base).
			L10.ParagraphFormat.LineSpacing = 10 * 1.5f

			' Set font size to 9 points for the second level of the first list style.
			L11.CharacterFormat.FontSize = 9

			' Enable legal-style numbering (e.g., 1.1, 1.2) for the second level of the first style.
			L11.IsLegalStyleNumbering = True

			' Use Arabic numerals (1, 2, 3...) as the numbering pattern.
			L11.PatternType = ListPatternType.Arabic

			' Specify that no character (like a dot or parenthesis) follows the number.
			L11.FollowCharacter = FollowCharacterType.Nothing

			' Set the bullet character using Unicode (U+006E, which is lowercase 'n'¡ªoften used as a custom bullet).
			L11.BulletCharacter = ChrW(&H006e).ToString()

			' Align the list number to the left within its numbering space.
			L11.NumberAlignment = ListNumberAlignment.Left

			' Position the number 10 points to the left of the default position (negative value).
			L11.NumberPosition = -10

			' Set the distance between the number and the following text tab stop to 0.5 inches.
			L11.TabSpaceAfter = 0.5f

			' Define the starting horizontal position of the actual list text as 0.5 inches.
			L11.TextPosition = 0.5f

			' Start numbering at 4 instead of 1 for this level.
			L11.StartAt = 4

			' Prefix each number with the word "Chapter" (e.g., "Chapter4").
			L11.NumberPrefix = "Chapter"

			' Allow numbering to restart when a higher-level item changes.
			L11.NoRestartByHigher = False

			' Do not inherit the numbering format from the previous (higher) level.
			L11.UsePrevLevelPattern = False

			' Set the font name to "Arial" for the third level of the first list style.
			L12.CharacterFormat.FontName = "Arial"

			' Set identical line spacing for the first level of the second list style.
			L20.ParagraphFormat.LineSpacing = 10 * 1.5f

			' Set font size to 9 points for the second level of the second list style.
			L21.CharacterFormat.FontSize = 9

			' Enable legal-style numbering for the second level of the second style.
			L21.IsLegalStyleNumbering = True

			' Use Arabic numerals for numbering in the second style¡¯s second level.
			L21.PatternType = ListPatternType.Arabic

			' No character follows the number in the second style¡¯s second level.
			L21.FollowCharacter = FollowCharacterType.Nothing

			' Use the same Unicode bullet character as in the first style.
			L21.BulletCharacter = ChrW(&H006e).ToString()

			' Left-align the number in the second style¡¯s second level.
			L21.NumberAlignment = ListNumberAlignment.Left

			' Position the number at the same horizontal offset (-10 points).
			L21.NumberPosition = -10

			' Set the same tab space after the number (0.5 inches).
			L21.TabSpaceAfter = 0.5f

			' Set the same text start position (0.5 inches).
			L21.TextPosition = 0.5f

			' Start numbering at 4 for consistency.
			L21.StartAt = 4

			' Apply the same "Chapter" prefix.
			L21.NumberPrefix = "Chapter"

			' Allow restart by higher levels.
			L21.NoRestartByHigher = False

			' Do not use the previous level¡¯s pattern.
			L21.UsePrevLevelPattern = False

			' Create a picture-based bullet for this level (this makes it different from L11).
			L21.CreatePictureBullet()

			' Set the font name to "Arial" for the third level of the second list style.
			L22.CharacterFormat.FontName = "Arial"

			' Compare the first levels of both list styles for equality.
			Dim r0 As Boolean = L10.Equals(L20)

			' Compare the second levels; note: L21 has a picture bullet while L11 does not, so they differ.
			Dim r1 As Boolean = L11.Equals(L21)

			' Compare the third levels, which only differ in their parent style but have identical properties.
			Dim r2 As Boolean = L12.Equals(L22)

			' Add a new section to the document to hold output paragraphs.
			Dim section As Section = document.AddSection()

			' Add a paragraph describing the first comparison.
			Dim paragraph As Paragraph = section.AddParagraph()

			' Set the text explaining what is being compared (level 1 vs level 1).
			paragraph.Text = "Compare the first level of the first ListStyle with the first level of the second ListStyle."

			' Add a new paragraph to display the result of the first comparison.
			paragraph = section.AddParagraph()

			' Insert the boolean result (true/false) into the paragraph text.
			paragraph.Text = "The comparison result is " & r0 & "."

			' Add a paragraph describing the second comparison.
			paragraph = section.AddParagraph()

			' Set the text for comparing the second levels.
			paragraph.Text = "Compare the second level of the first ListStyle with the second level of the second ListStyle."

			' Add a new paragraph for the result.
			paragraph = section.AddParagraph()

			' Display the result of the second comparison (expected: false due to picture bullet).
			paragraph.Text = "The comparison result is " & r1 & "."

			' Add a paragraph describing the third comparison.
			paragraph = section.AddParagraph()

			' Set the text for comparing the third levels.
			paragraph.Text = "Compare the third level of the first ListStyle with the third level of the second ListStyle."

			' Add a new paragraph for the final result.
			paragraph = section.AddParagraph()

			' Display the result of the third comparison (expected: true if only font name is set identically).
			paragraph.Text = "The comparison result is " & r2 & "."

			' Save the document to a file named "result.docx" in Word 2016 format.
			document.SaveToFile("result.docx", FileFormat.Docx2016)

			' Close the document to release internal resources.
			document.Close()

			' Explicitly dispose of the document object to free memory.
			document.Dispose()

			'Launching the Word file.
			WordDocViewer("result.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
