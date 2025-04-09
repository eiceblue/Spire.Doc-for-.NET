Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports System.IO

Namespace GetText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object and specify the path of the Word document to extract text from.
			Dim document As New Document("..\..\..\..\..\..\Data\ExtractText.docx")

			' Retrieve the text content from the Document object and store it in a string variable.
			Dim text As String = document.GetText()

			' Write the extracted text to a text file named "Extract.txt".
			File.WriteAllText("Extract.txt", text)

			' Dispose of the Document object to release any resources associated with it.
			document.Dispose()

			'launch the file.
			WordDocViewer("Extract.txt")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
