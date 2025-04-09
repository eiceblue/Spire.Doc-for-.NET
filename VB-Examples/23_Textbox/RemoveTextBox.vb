Imports Spire.Doc

Namespace RemoveTextBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Specify the input file path
			Dim input As String = "..\..\..\..\..\..\Data\TextBoxTemplate.docx"

			' Create a new instance of Document
			Dim doc As New Document()

			' Load the document from the specified input file
			doc.LoadFromFile(input)

			' Remove the first text box in the document
			doc.TextBoxes.RemoveAt(0)

			' Clear all the text boxes in the document
			'doc.TextBoxes.Clear();

			' Specify the output file path
			Dim output As String = "RemoveTextBox.docx"

			' Save the modified document to the output file with the specified file format (Docx)
			doc.SaveToFile(output, FileFormat.Docx)

			' Dispose the document object to release resources
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
