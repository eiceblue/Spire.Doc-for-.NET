Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports System.IO

Namespace CountVariables
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

			' Get the number of variables in the document
			Dim number As Integer = document.Variables.Count

			' Initialize a StringBuilder to store the content
			Dim content As New StringBuilder()
			content.AppendLine("The number of variables is: " & number.ToString())

			' Specify the file path for the output result
			Dim result As String = "Result-CountVariables.txt"

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
