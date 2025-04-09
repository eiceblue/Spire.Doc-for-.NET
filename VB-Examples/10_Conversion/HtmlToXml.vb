Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents

Namespace HtmlToXml
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create Word document.
			Dim document As New Document()

			'Load the file from disk.
			document.LoadFromFile("..\..\..\..\..\..\Data\Template_HtmlFile.html")

			Dim result As String = "Result-HtmlToXml.xml"

			'Save to file.
			document.SaveToFile(result, FileFormat.Xml)

			'Dispose the document
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
