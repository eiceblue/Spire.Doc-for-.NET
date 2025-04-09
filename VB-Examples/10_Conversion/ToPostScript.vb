Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc

Namespace ToPostScript
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document
			Dim doc As New Document()

			'Load the file from disk
			doc.LoadFromFile("..\..\..\..\..\..\Data\ConvertedTemplate.docx")

			Dim result As String = "ToPostScript.ps"
			
			'Save to file
			doc.SaveToFile(result, FileFormat.PostScript)

			'Dispose the document
			doc.Dispose()

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
