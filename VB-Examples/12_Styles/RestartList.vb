Imports Spire.Doc
Imports Spire.Doc.Collections
Imports Spire.Doc.Documents

Namespace RestartList
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
			Dim document As New Document()

			'Create a new section
			Dim section As Section = document.AddSection()

			'Create a new paragraph
			Dim paragraph As Paragraph = section.AddParagraph()

			'Append Text
			paragraph.AppendText("List 1")


			Dim numberList As ListStyle = document.Styles.Add(ListType.Numbered, "Numbered1")


			'Add paragraph and apply the list style
			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 1")
			paragraph.ListFormat.ApplyStyle(numberList.Name)

			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 2")
			paragraph.ListFormat.ApplyStyle(numberList.Name)

			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 3")
			paragraph.ListFormat.ApplyStyle(numberList.Name)

			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 4")
			paragraph.ListFormat.ApplyStyle(numberList.Name)

			'Append Text
			paragraph = section.AddParagraph()
			paragraph.AppendText("List 2")


			Dim numberList2 As ListStyle = document.Styles.Add(ListType.Numbered, "Numbered2")
			Dim Levels As ListLevelCollection = numberList2.ListRef.Levels

			'set start number of second list
			Levels(0).StartAt = 10


			'Add paragraph and apply the list style
			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 5")
			paragraph.ListFormat.ApplyStyle(numberList2.Name)

			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 6")
			paragraph.ListFormat.ApplyStyle(numberList2.Name)

			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 7")
			paragraph.ListFormat.ApplyStyle(numberList2.Name)

			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 8")
			paragraph.ListFormat.ApplyStyle(numberList2.Name)

			'Save to docx file.
			Dim output As String = "RestartList.docx"
			document.SaveToFile(output, FileFormat.Docx)

			'Dispose the document
			document.Dispose()
			WordDocViewer(output)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
