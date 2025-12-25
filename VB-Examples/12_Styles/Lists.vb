Imports Spire.Doc
Imports Spire.Doc.Collections
Imports Spire.Doc.Documents


Namespace Text
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create a Word document
			Dim document As New Document()

			'Add a section
			Dim sec As Spire.Doc.Section = document.AddSection()

			'Add paragraph and apply style
			Dim paragraph As Spire.Doc.Documents.Paragraph = sec.AddParagraph()
			paragraph.AppendText("Lists")
			paragraph.ApplyStyle(BuiltinStyle.Title)
			paragraph = sec.AddParagraph()
			paragraph.AppendText("Numbered List:").CharacterFormat.Bold = True

			'Create list style
			Dim listStyle As ListStyle = document.Styles.Add(ListType.Numbered, "numberList")
			Dim Levels As ListLevelCollection = listStyle.ListRef.Levels
			Levels(1).NumberPrefix = ChrW(&H0000).ToString() & "."
			Levels(1).PatternType = ListPatternType.Arabic
			Levels(2).NumberPrefix = ChrW(&H0000).ToString() & "." & ChrW(&H0001).ToString() & "."
			Levels(2).PatternType = ListPatternType.Arabic

			Dim bulletList As ListStyle = document.Styles.Add(ListType.Bulleted, "bulletList")
			'Add paragraph and apply the list style
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 1")
			paragraph.ListFormat.ApplyStyle(listStyle)

			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2")
			paragraph.ListFormat.ApplyStyle(listStyle)

			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.1")
			paragraph.ListFormat.ApplyStyle(listStyle)
			paragraph.ListFormat.ListLevelNumber = 1

			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.2")
			paragraph.ListFormat.ApplyStyle(listStyle)
			paragraph.ListFormat.ListLevelNumber = 1

			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.2.1")
			paragraph.ListFormat.ApplyStyle(listStyle)
			paragraph.ListFormat.ListLevelNumber = 2
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.2.2")
			paragraph.ListFormat.ApplyStyle(listStyle)
			paragraph.ListFormat.ListLevelNumber = 2
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.2.3")
			paragraph.ListFormat.ApplyStyle(listStyle)
			paragraph.ListFormat.ListLevelNumber = 2

			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.3")
			paragraph.ListFormat.ApplyStyle(listStyle)
			paragraph.ListFormat.ListLevelNumber = 1

			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 3")
			paragraph.ListFormat.ApplyStyle(listStyle)

			paragraph = sec.AddParagraph()
			paragraph.AppendText("Bulleted List:").CharacterFormat.Bold = True

			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 1")
			paragraph.ListFormat.ApplyStyle(bulletList)
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2")
			paragraph.ListFormat.ApplyStyle(bulletList)

			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.1")
			paragraph.ListFormat.ApplyStyle(bulletList)
			paragraph.ListFormat.ListLevelNumber = 1
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 2.2")
			paragraph.ListFormat.ApplyStyle(bulletList)
			paragraph.ListFormat.ListLevelNumber = 1
			paragraph = sec.AddParagraph()
			paragraph.AppendText("List Item 3")
			paragraph.ListFormat.ApplyStyle(bulletList)



			'Save doc file.
			document.SaveToFile("lists-out.docx", FileFormat.Docx)
			document.Close()

		   'Launching the MS Word file.
		   WordDocViewer("lists-out.docx")


		End Sub
		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
