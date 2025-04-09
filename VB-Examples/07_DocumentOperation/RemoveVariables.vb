Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace RemoveVariables
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

			' Remove the variable with the specified name from the document
			document.Variables.Remove("A1")

			' Set the IsUpdateFields property to true to update fields when saving the document
			document.IsUpdateFields = True

			' Specify the file path for the output result
			Dim result As String = "Result-RemoveVariables.docx"

			' Save the document to a file with the specified file format
			document.SaveToFile(result, FileFormat.Docx2013)

			' Dispose of the document object
			document.Dispose()

			'Launch the MS Word file.
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
