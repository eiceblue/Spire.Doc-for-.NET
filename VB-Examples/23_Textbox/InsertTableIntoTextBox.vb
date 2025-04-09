Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace InsertTableIntoTextBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new document
			Dim doc As New Document()

			' Add a section to the document
			Dim section As Section = doc.AddSection()

			' Add a paragraph to the section
			Dim paragraph As Paragraph = section.AddParagraph()

			' Append a text box to the paragraph with specified dimensions
			Dim textbox As Spire.Doc.Fields.TextBox = paragraph.AppendTextBox(300, 100)

			' Set the horizontal and vertical positioning of the text box
			textbox.Format.HorizontalOrigin = HorizontalOrigin.Page
			textbox.Format.HorizontalPosition = 140
			textbox.Format.VerticalOrigin = VerticalOrigin.Page
			textbox.Format.VerticalPosition = 50

			' Add a paragraph to the text box
			Dim textboxParagraph As Paragraph = textbox.Body.AddParagraph()

			' Append text to the paragraph in the text box
			Dim textboxRange As TextRange = textboxParagraph.AppendText("Table 1")
			textboxRange.CharacterFormat.FontName = "Arial"

			' Add a table to the body of the text box
			Dim table As Table = textbox.Body.AddTable(True)

			' Reset the number of rows and columns in the table
			table.ResetCells(4, 4)

			' Define the data for the table
			Dim data(,) As String = { {"Name","Age","Gender","ID" }, {"John","28","Male","0023" }, {"Steve","30","Male","0024" }, {"Lucy","26","female","0025" } }

			' Populate the table with data
			For i As Integer = 0 To 3
				For j As Integer = 0 To 3
					Dim tableRange As TextRange = table(i, j).AddParagraph().AppendText(data(i, j))
					tableRange.CharacterFormat.FontName = "Arial"
				Next j
			Next i

			' Apply a predefined table style to the table
			table.ApplyStyle(DefaultTableStyle.TableColorful2)

			' Save the document to a file
			Dim output As String = "InsertTableIntoTextBox.docx"
			doc.SaveToFile(output, FileFormat.Docx)

			' Dispose the document object
			doc.Dispose()
			
			Viewer(output)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
