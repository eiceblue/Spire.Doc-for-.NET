Imports Spire.Doc
Imports Spire.Doc.Collections
Imports Spire.Doc.Documents

Namespace CreatePictureBullet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object to represent a Word document.
			Dim document As New Document()

			' Add a new section to the document. A section is a fundamental structural unit in Word.
			Dim sec As Section = document.AddSection()

			' Add a paragraph to the section and assign it to the variable 'paragraph'.
			Dim paragraph As Spire.Doc.Documents.Paragraph = sec.AddParagraph()

			' Create a bulleted list style named "bulletList".
			Dim listStyle As ListStyle = document.Styles.Add(ListType.Bulleted, "bulletList")

			' Get the collection of list levels (up to 9 levels, indexed from 0) for this list style.
			Dim Levels As ListLevelCollection = listStyle.ListRef.Levels

			' Enable picture bullet for level 0 (the top-level list items).
			Levels(0).CreatePictureBullet()

			' Load an image file as the bullet for level 0 using a relative path.
			Levels(0).PictureBullet.LoadImage("..\..\..\..\..\..\Data\Word.png")

			' Enable picture bullet for level 1 (second-level indented list items).
			Levels(1).CreatePictureBullet()

			' Load another image file as the bullet for level 1 using a relative path.
			Levels(1).PictureBullet.LoadImage("..\..\..\..\..\..\Data\logo.png")

			' Add a new paragraph for the first-level list item.
			paragraph = sec.AddParagraph()

			' Append the text "List Item 1" to the paragraph.
			paragraph.AppendText("List Item 1")

			' Apply the custom "bulletList" style to this paragraph.
			paragraph.ListFormat.ApplyStyle(listStyle)

			' Add another paragraph for the second-level list item.
			paragraph = sec.AddParagraph()

			' Append the text "List Item 1.1" to this paragraph.
			paragraph.AppendText("List Item 1.1")

			' Apply the same list style to make it part of the list structure.
			paragraph.ListFormat.ApplyStyle(listStyle)

			' Set the list level of this paragraph to 1 (i.e., second indentation level).
			paragraph.ListFormat.ListLevelNumber = 1

			' Save the document as "CreatePictureBullet.docx" in Word 2016 (.docx) format.
			document.SaveToFile("CreatePictureBullet.docx", FileFormat.Docx2016)

			' Close the document to release internal resources such as file handles.
			document.Close()

			' Explicitly dispose of the document object to ensure memory is freed immediately.
			document.Dispose()


			'Launching the Word file.
			WordDocViewer("CreatePictureBullet.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
