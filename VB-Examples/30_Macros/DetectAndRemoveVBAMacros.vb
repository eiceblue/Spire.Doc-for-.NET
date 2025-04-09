Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace DetectAndRemoveVBAMacros
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			' Create a new instance of Document
			Dim document As New Document()

			' Load the Word document from the specified file that may contain VBA macros
			document.LoadFromFile("..\..\..\..\..\..\Data\DetectAndRemoveVBAMacros.docm")

			' Check if the document contains VBA macros
			If document.IsContainMacro Then
				' Clear/remove the VBA macros from the document
				document.ClearMacros()
			End If

			' Specify the name for the resulting document after removing VBA macros
			Dim result As String = "Result-DetectAndRemoveVBAMacros.docm"

			' Save the modified document to a new file with the specified name and format (Docm for macro-enabled document)
			document.SaveToFile(result, FileFormat.Docm)

			' Dispose of the document object when finished using it
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
