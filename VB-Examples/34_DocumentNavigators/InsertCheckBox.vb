Imports Spire.Doc
Imports Spire.Doc.Documents


Namespace InsertCheckBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object to represent a Word document.
			Dim doc As New Document()

			' Initialize a DocumentNavigator to assist with content insertion and navigation.
			Dim navigator As New DocumentNavigator(doc)

			' Get the current section of the document where content will be added.
			Dim section As Section = navigator.CurrentSection

			' Add a new paragraph to the section.
			Dim paragraph As Paragraph = section.AddParagraph()

			' Append descriptive text indicating a checkbox control will follow (default unselected).
			paragraph.AppendText("Add CheckBox Content Control: ")

			' Append additional label text for the first checkbox example.
			paragraph.AppendText("Default: Unselected: ")

			' Move the cursor to the beginning of the first paragraph (index 1, offset 0) to insert content precisely.
			navigator.MoveToParagraph(1, 0)

			' Insert an unchecked checkbox content control named "Checkbox1" with a size of 20 points.
			navigator.InsertCheckBox("Checkbox1", False, 20)

			' Add a new paragraph for the second checkbox example.
			paragraph = section.AddParagraph()

			' Append label text indicating this checkbox is selected by default.
			paragraph.AppendText("Default: Selected: ")

			' Move the cursor to the beginning of the second paragraph (index 2, offset 0).
			navigator.MoveToParagraph(2, 0)

			' Insert a checked checkbox content control named "Checkbox2" with a size of 20 points.
			navigator.InsertCheckBox("Checkbox2", True, 20)

			' Add a new paragraph for the third checkbox example.
			paragraph = section.AddParagraph()

			' Append label text explaining that this checkbox starts unchecked but will be set to checked via API.
			paragraph.AppendText("Default: Uncheck. Set to check now: ")

			' Move the cursor to the beginning of the third paragraph (index 3, offset 0).
			navigator.MoveToParagraph(3, 0)

			' Insert a checkbox named "Checkbox3" that is initially unchecked (false), then programmatically marked as checked (true), with size 20.
			navigator.InsertCheckBox("Checkbox3", False, True, 20)

			' Add a new paragraph for the fourth checkbox example.
			paragraph = section.AddParagraph()

			' Append label text for a checkbox that appears selected by default in the UI but is actually stored as unchecked.
			paragraph.AppendText("Default selection, unselected by default: ")

			' Move the cursor to the beginning of the fourth paragraph (index 4, offset 0).
			navigator.MoveToParagraph(4, 0)

			' Insert a checkbox named "Checkbox4" that appears checked (true) but is stored as unchecked (false), with size 20.
			navigator.InsertCheckBox("Checkbox4", True, False, 20)

			' Save the resulting document to a file named "InsertCheckBox.docx" in DOCX format.
			doc.SaveToFile("InsertCheckBox.docx", FileFormat.Docx)


			' Close the document to release internal resources.
			doc.Close()

			' Explicitly dispose of the document object to free memory immediately.
			doc.Dispose()

			'Launch the MS Word file.
			WordDocViewer("InsertCheckBox.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
