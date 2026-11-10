Imports Spire.Doc
Imports Spire.Doc.Collections
Imports Spire.Doc.Documents
Imports Spire.Doc.Interface

Namespace ModifyCustomStyles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document instance.
			Dim document As New Document()

			' Load an existing Word document that contains custom styles from the specified relative file path.
			document.LoadFromFile("..\..\..\..\..\..\..\Data\CustomStyles.docx")

			' Attempt to find a paragraph style named "MyParagraphStyle1" in the document's style collection.
			Dim style1 As IStyle = document.Styles.FindByName("MyParagraphStyle1")

			' Check if the style exists and is of type ParagraphStyle.
			If style1 IsNot Nothing AndAlso TypeOf style1 Is ParagraphStyle Then
				' Cast the found style to ParagraphStyle for modification.
				Dim cStyle As ParagraphStyle = TryCast(style1, ParagraphStyle)

				' Set the font name of the paragraph style to "Arial".
				cStyle.CharacterFormat.FontName = "Arial"

				' Set the font size to 18 points.
				cStyle.CharacterFormat.FontSize = 18

				' Disable bold formatting.
				cStyle.CharacterFormat.Bold = False

				' Change the text color to gray.
				cStyle.CharacterFormat.TextColor = Color.Gray
			End If

			' Attempt to find a table style named "MyTableStyle1" in the document's style collection.
			Dim style2 As IStyle = document.Styles.FindByName("MyTableStyle1")

			' Check if the style exists and is of type TableStyle.
			If style2 IsNot Nothing AndAlso TypeOf style2 Is TableStyle Then
				' Cast the found style to TableStyle for modification.
				Dim tableStyle As TableStyle = TryCast(style2, TableStyle)

				' Set the border color of the table to gray.
				tableStyle.Borders.Color = Color.Gray

				' Change the border type to single line.
				tableStyle.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single

				' Set the border line width to 1 point.
				tableStyle.Borders.LineWidth = 1

				' Apply row striping every 2 rows (i.e., stripe pattern repeats every two rows).
				tableStyle.RowStripe = 2
			End If

			' Save the modified document to a new file using the Word 2016 .docx format.
			document.SaveToFile("ModifyCustomStyles.docx", FileFormat.Docx2016)

			' Close the document to release associated resources.
			document.Close()

			' Explicitly dispose of the document object to free memory immediately.
			document.Dispose()

			'Launching the Word file.
			WordDocViewer("ModifyCustomStyles.docx")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
