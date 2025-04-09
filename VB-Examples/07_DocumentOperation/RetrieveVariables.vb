Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.IO

Namespace RetrieveVariables
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of the Document class
			Dim document As New Document()

			' Load the specified document file
			document.LoadFromFile("..\..\..\..\..\..\..\Data\Template_Docx_6.docx")

			' Retrieve the name of the variable at index 0
			Dim s1 As String = document.Variables.GetNameByIndex(0)

			' Retrieve the value of the variable at index 0
			Dim s2 As String = document.Variables.GetValueByIndex(0)

			' Retrieve the value of the variable with the specified name
			Dim s3 As String = document.Variables("A1")

			' Initialize a StringBuilder to store the content
			Dim content As New StringBuilder()
			content.AppendLine("The name of the variable retrieved by index 0 is: " & s1)
			content.AppendLine("The value of the variable retrieved by index 0 is: " & s2)
			content.AppendLine("The value of the variable retrieved by name ""A1"" is: " & s3)

			' Specify the file path for the output result
			Dim result As String = "Result-RetrieveVariables.txt"

			' Write the content to a text file
			File.WriteAllText(result, content.ToString())

			' Dispose of the document object
			document.Dispose()


			'Launch the file.
			WordDocViewer(result)
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
