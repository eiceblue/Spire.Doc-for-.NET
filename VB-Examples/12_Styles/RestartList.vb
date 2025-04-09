Imports Spire.Doc
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

			Dim numberList As New ListStyle(document, ListType.Numbered)
			numberList.Name = "Numbered1"
			document.ListStyles.Add(numberList)

			'Add paragraph and apply the list style
			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 1")
			paragraph.ListFormat.ApplyStyle(numberList.Name)

			'Add paragraph and apply the list style
			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 2")
			paragraph.ListFormat.ApplyStyle(numberList.Name)

			'Add paragraph and apply the list style
			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 3")
			paragraph.ListFormat.ApplyStyle(numberList.Name)

			'Add paragraph and apply the list style
			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 4")
			paragraph.ListFormat.ApplyStyle(numberList.Name)

			'Add paragraph and append text
			paragraph = section.AddParagraph()
			paragraph.AppendText("List 2")

			'Create a numbered list
			Dim numberList2 As New ListStyle(document, ListType.Numbered)
			numberList2.Name = "Numbered2"
			'set start number of second list
			numberList2.Levels(0).StartAt = 10
			document.ListStyles.Add(numberList2)

			'Add paragraph and apply the numbered list
			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 5")
			paragraph.ListFormat.ApplyStyle(numberList2.Name)

			'Add paragraph and apply the numbered list
			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 6")
			paragraph.ListFormat.ApplyStyle(numberList2.Name)

			'Add paragraph and apply the numbered list
			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 7")
			paragraph.ListFormat.ApplyStyle(numberList2.Name)

			'Add paragraph and apply the numbered list
			paragraph = section.AddParagraph()
			paragraph.AppendText("List Item 8")
			paragraph.ListFormat.ApplyStyle(numberList2.Name)

			'Save the document
			Dim output As String = "RestartList.docx"
			document.SaveToFile(output)

			'Dispose of the document object
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
