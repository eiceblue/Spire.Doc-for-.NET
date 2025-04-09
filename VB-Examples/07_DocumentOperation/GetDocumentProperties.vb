Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.IO
Imports System.Text

Namespace GetDocumentProperties
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new Document object
			Dim document As New Document()

			' Load the document from the specified file
			document.LoadFromFile("..\..\..\..\..\..\Data\Properties.docx")

			' Create a new StringBuilder to store the content
			Dim content As New StringBuilder()

			' Get the built-in document properties and store their values
			Dim title As String = document.BuiltinDocumentProperties.Title
			Dim comments As String = document.BuiltinDocumentProperties.Comments
			Dim author As String = document.BuiltinDocumentProperties.Author
			Dim keywords As String = document.BuiltinDocumentProperties.Keywords
			Dim company As String = document.BuiltinDocumentProperties.Company

			' Create a result string with the built-in document properties values
			Dim result As String = String.Format("The Builtin document properties:" & vbCrLf & "Title: " & title & "." & vbCrLf & "Comments: " & comments & "." & vbCrLf & "Author: " & author & "." & vbCrLf & "Keywords: " & keywords & "." & vbCrLf & "Company: " & company)

			' Append the result to the content with a new line
			content.AppendLine(result & vbCrLf & "The custom document properties:")

			' Iterate through the custom document properties and append their names and values to the content
			For i As Integer = 0 To document.CustomDocumentProperties.Count - 1
				Dim customProperties As String = String.Format(document.CustomDocumentProperties(i).Name & ": " & document.CustomDocumentProperties(i).Value)
				content.AppendLine(customProperties)
			Next i

			' Write the content to a text file
			File.WriteAllText("Output.txt", content.ToString())

			' Release all resources used by the Document object
			document.Dispose()


			'Launch the txt file.
			WordDocViewer("Output.txt")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
