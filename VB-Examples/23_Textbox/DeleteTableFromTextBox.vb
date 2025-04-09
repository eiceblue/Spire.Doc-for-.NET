Imports System.Text
Imports System.IO
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace DeleteTableFromTextBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path
			Dim input As String = "..\..\..\..\..\..\Data\TextBoxTable.docx"

			' Create a new Document object
			Dim doc As New Document()

			' Load a Word document from the specified input file
			doc.LoadFromFile(input)

			' Access the first text box in the document
			Dim textbox As Spire.Doc.Fields.TextBox = doc.TextBoxes(0)

			' Remove the table inside the text box
			textbox.Body.Tables.RemoveAt(0)

			' Specify the output file name
			Dim output As String = "DeleteTableFromTextBox.docx"

			' Save the modified document to a new file
			doc.SaveToFile(output, FileFormat.Docx)

			' Dispose the document object to free up resources
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
