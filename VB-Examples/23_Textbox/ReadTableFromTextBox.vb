Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace ReadTableFromTextBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Load the Word document from the specified file path
			Dim input As String = "..\..\..\..\..\..\Data\TextBoxTable.docx"
			Dim doc As New Document()
			doc.LoadFromFile(input)

			' Get the first textbox in the document
			Dim textbox As Spire.Doc.Fields.TextBox = doc.TextBoxes(0)

			' Get the first table from the textbox
			Dim table As Table = TryCast(textbox.Body.Tables(0), Table)

			' Initialize an empty string to store the table data
			Dim str As String = Nothing

			' Iterate through each row in the table
			For Each row As TableRow In table.Rows
				' Iterate through each cell in the row
				For Each cell As TableCell In row.Cells
					' Iterate through each paragraph in the cell
					For Each paragraph As Paragraph In cell.Paragraphs
						' Append the text of each paragraph to the string, separated by a tab
						str &= paragraph.Text & vbTab
					Next paragraph
				Next cell
				' Add a new line after processing each row
				str &= vbCrLf
			Next row

			' Specify the output file path
			Dim output As String = "ReadTableFromTextBox.txt"

			' Write the table data to the output file
			File.WriteAllText(output, str)

			' Dispose of the document object
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
