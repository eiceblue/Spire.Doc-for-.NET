Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace Text
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a New document
			Dim document As New Document()

			'Add a section
			Dim sec As Section = document.AddSection()

			'Add a paragraph
			Dim paragraph As Paragraph = sec.AddParagraph()

			'Append text
			paragraph.AppendText("Lists")

			'Apply the builtin style
			paragraph.ApplyStyle(BuiltinStyle.Title)

			'Add a new paragraph
			paragraph = sec.AddParagraph()

			'Append text and set bold style
			paragraph.AppendText("Numbered List:").CharacterFormat.Bold = True


			'Create a new numbered list
			Dim numberList As New ListStyle(document, ListType.Numbered)
			numberList.Name = "numberList"
			numberList.Levels(1).NumberPrefix = ChrW(&H0).ToString() & "."
			numberList.Levels(1).PatternType = ListPatternType.Arabic
			numberList.Levels(2).NumberPrefix = ChrW(&H0).ToString() & "." & ChrW(&H1).ToString() & "."
			numberList.Levels(2).PatternType = ListPatternType.Arabic

			'Create a bulleted list
			Dim bulletList As New ListStyle(document, ListType.Bulleted)
			bulletList.Name = "bulletList"

			'Add the two list style to the document
			document.ListStyles.Add(numberList)
			document.ListStyles.Add(bulletList)

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 1")
			paragraph.ListFormat.ApplyStyle(numberList.Name)

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2")
			paragraph.ListFormat.ApplyStyle(numberList.Name)

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.1")
			paragraph.ListFormat.ApplyStyle(numberList.Name)
			paragraph.ListFormat.ListLevelNumber = 1

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.2")
			paragraph.ListFormat.ApplyStyle(numberList.Name)
			paragraph.ListFormat.ListLevelNumber = 1

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.2.1")
			paragraph.ListFormat.ApplyStyle(numberList.Name)
			paragraph.ListFormat.ListLevelNumber = 2

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.2.2")
			paragraph.ListFormat.ApplyStyle(numberList.Name)
			paragraph.ListFormat.ListLevelNumber = 2

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.2.3")
			paragraph.ListFormat.ApplyStyle(numberList.Name)
			paragraph.ListFormat.ListLevelNumber = 2

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.3")
			paragraph.ListFormat.ApplyStyle(numberList.Name)
			paragraph.ListFormat.ListLevelNumber = 1

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 3")
			paragraph.ListFormat.ApplyStyle(numberList.Name)

			'Add a paragraph and apply bold style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Bulleted List:").CharacterFormat.Bold = True

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 1")
			paragraph.ListFormat.ApplyStyle(bulletList.Name)

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2")
			paragraph.ListFormat.ApplyStyle(bulletList.Name)

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.1")
			paragraph.ListFormat.ApplyStyle(bulletList.Name)
			paragraph.ListFormat.ListLevelNumber = 1

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.2")
			paragraph.ListFormat.ApplyStyle(bulletList.Name)
			paragraph.ListFormat.ListLevelNumber = 1

			'Add a paragraph and apply style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 3")
			paragraph.ListFormat.ApplyStyle(bulletList.Name)

			'Save the document
			Dim filePath As String = "Lists.docx"
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
