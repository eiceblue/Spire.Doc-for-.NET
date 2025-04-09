Imports System.ComponentModel
Imports System.Text
Imports Spire.Doc
Imports Spire.Doc.Documents
Imports Spire.Doc.Fields

Namespace ToXML
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			'Create word document.
			Dim document As New Document()

			'Load the file from disk
			document.LoadFromFile("..\..\..\..\..\..\Data\Summary_of_Science.doc")

			'Save the document to a xml file.
			document.SaveToFile("Sample.xml", FileFormat.Xml)

			'Dispose the document
			document.Dispose()

			'Launch the file.
			XMLViewer("Sample.xml")
		End Sub
		Private Sub XMLViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

	End Class
End Namespace
