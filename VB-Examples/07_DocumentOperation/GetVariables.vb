Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace GetVariables
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

			' Initialize a StringBuilder to store the content
			Dim stringBuilder As New StringBuilder()

			' Add initial message to the StringBuilder
			stringBuilder.AppendLine("This document has following variables:")

			' Iterate through each variable in the document and append their name and value to the StringBuilder
			For Each entry As KeyValuePair(Of String, String) In document.Variables
				Dim name As String = entry.Key
				Dim value As String = entry.Value
				stringBuilder.AppendLine("Name: " & name & ", " & "Value: " & value)
			Next entry

			' Write the content of the StringBuilder to a text file
			File.WriteAllText("GetVariables_out.txt", stringBuilder.ToString())

			' Dispose of the document object
			document.Dispose()

			WordDocViewer("GetVariables_out.txt")
		End Sub

		Private Sub WordDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub


	End Class
End Namespace
